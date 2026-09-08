"""Guards for input that arrives from the public internet.

The engine was written for files its author wrote. Two of its behaviours are
harmless in that setting and dangerous once anyone can upload:

* `Image Src` is fetched with `urlopen`. On a box that also runs Postgres, an
  internal wiki, a Tailscale interface and a cloud metadata endpoint, that is a
  server-side request forgery primitive handed to the uploader.
* An `Image Src` that is not a URL is opened as a filesystem path.

`fetch_remote_image` closes the first. The second is closed by
`catalog_exporter.configure_image_loading(allow_local_paths=False)`.
"""

from __future__ import annotations

import http.client
import ipaddress
import socket
import ssl
from urllib.parse import urljoin, urlsplit

USER_AGENT = "topdf.stuffs.bid catalog exporter (+https://stuffs.bid/topdf)"

MAX_IMAGE_BYTES = 12 * 1024 * 1024
MAX_DOCUMENT_BYTES = 4 * 1024 * 1024
FETCH_TIMEOUT_SECONDS = 12
MAX_REDIRECTS = 3
ALLOWED_PORTS = {80, 443}

# `ipaddress`'s own `is_private` misses the shared CGNAT range, which is exactly
# where this machine's Tailscale addresses live -- so it is listed explicitly
# rather than trusted to be covered.
EXTRA_BLOCKED_NETWORKS = [
    ipaddress.ip_network("100.64.0.0/10"),   # RFC 6598 shared / Tailscale
    ipaddress.ip_network("192.0.0.0/24"),    # IETF protocol assignments
    ipaddress.ip_network("198.18.0.0/15"),   # benchmarking
    ipaddress.ip_network("::ffff:0:0/96"),   # IPv4-mapped IPv6
    ipaddress.ip_network("64:ff9b::/96"),    # NAT64
]


class ImageFetchError(Exception):
    """Raised when a product image cannot be fetched safely.

    The engine already treats a failed image as "render the card without a
    photo", so a blocked URL degrades the page rather than failing the export.
    """


def _reject_if_not_public(raw_ip: str) -> None:
    address = ipaddress.ip_address(raw_ip)
    if (
        address.is_private
        or address.is_loopback
        or address.is_link_local
        or address.is_reserved
        or address.is_multicast
        or address.is_unspecified
        or not address.is_global
    ):
        raise ImageFetchError(f"address {raw_ip} is not a public address")
    for network in EXTRA_BLOCKED_NETWORKS:
        if address.version == network.version and address in network:
            raise ImageFetchError(f"address {raw_ip} is in a blocked range")


def _resolve_public_address(host: str, port: int) -> tuple[str, int]:
    """Resolve `host`, rejecting the name outright if *any* answer is internal.

    Checking every answer and then connecting to the one that was checked is
    what makes DNS rebinding pointless here: the second lookup that a rebinding
    attack depends on never happens.
    """
    try:
        candidates = socket.getaddrinfo(host, port, proto=socket.IPPROTO_TCP)
    except socket.gaierror as error:
        raise ImageFetchError(f"cannot resolve {host}") from error
    if not candidates:
        raise ImageFetchError(f"cannot resolve {host}")

    for candidate in candidates:
        _reject_if_not_public(candidate[4][0])

    family, _type, _proto, _canon, sockaddr = candidates[0]
    return sockaddr[0], family


def _request_once(url: str, accept: str, max_bytes: int) -> tuple[int, str | None, bytes]:
    parts = urlsplit(url)
    if parts.scheme not in ("http", "https"):
        raise ImageFetchError(f"unsupported scheme {parts.scheme!r}")
    if not parts.hostname:
        raise ImageFetchError("missing host")
    if parts.username or parts.password:
        raise ImageFetchError("credentials in URL are not accepted")

    secure = parts.scheme == "https"
    port = parts.port or (443 if secure else 80)
    if port not in ALLOWED_PORTS:
        raise ImageFetchError(f"port {port} is not allowed")

    address, family = _resolve_public_address(parts.hostname, port)

    sock = socket.socket(family, socket.SOCK_STREAM)
    sock.settimeout(FETCH_TIMEOUT_SECONDS)
    try:
        sock.connect((address, port))
        if secure:
            context = ssl.create_default_context()
            sock = context.wrap_socket(sock, server_hostname=parts.hostname)

        # The connection is already open to the address that was vetted, so
        # http.client only has to speak the protocol over it. Host is supplied
        # explicitly (which also stops http.client appending a port to it) so
        # the origin sees the name from the URL, not the pinned address.
        connection = http.client.HTTPConnection(parts.hostname, port, timeout=FETCH_TIMEOUT_SECONDS)
        connection.sock = sock
        target = parts.path or "/"
        if parts.query:
            target = f"{target}?{parts.query}"
        host_header = parts.hostname if port in (80, 443) else f"{parts.hostname}:{port}"
        connection.request(
            "GET",
            target,
            headers={"Host": host_header, "User-Agent": USER_AGENT, "Accept": accept},
        )
        response = connection.getresponse()

        if response.status in (301, 302, 303, 307, 308):
            return response.status, response.getheader("Location"), b""

        if response.status != 200:
            raise ImageFetchError(f"HTTP {response.status}")

        declared = response.getheader("Content-Length")
        if declared and declared.isdigit() and int(declared) > max_bytes:
            raise ImageFetchError("response is larger than the size limit")

        body = response.read(max_bytes + 1)
        if len(body) > max_bytes:
            raise ImageFetchError("response is larger than the size limit")
        if not body:
            raise ImageFetchError("empty response")
        return 200, None, body
    except ssl.SSLError as error:
        raise ImageFetchError(f"TLS error: {error}") from error
    except (socket.timeout, TimeoutError) as error:
        raise ImageFetchError("timed out") from error
    except OSError as error:
        raise ImageFetchError(f"connection failed: {error}") from error
    finally:
        try:
            sock.close()
        except OSError:
            pass


def fetch_guarded(url: str, *, accept: str, max_bytes: int) -> tuple[str, bytes]:
    """Fetch `url` through the guard, following redirects. Returns (final URL, body).

    The guard matters more here than it did for images. An image URL arrives
    inside a spreadsheet somebody assembled; a store URL is typed straight into
    a box on the page, which is as direct a server-side request forgery handle
    as this app has. Same resolve-then-connect vetting, one entry point.
    """
    current = url
    for _hop in range(MAX_REDIRECTS + 1):
        status, location, body = _request_once(current, accept, max_bytes)
        if status == 200:
            return current, body
        if not location:
            raise ImageFetchError(f"HTTP {status} without a redirect target")
        current = urljoin(current, location)
    raise ImageFetchError("too many redirects")


def fetch_remote_image(url: str) -> bytes:
    """Fetch a product image, refusing anything that points inside the network."""
    return fetch_guarded(url, accept="image/*", max_bytes=MAX_IMAGE_BYTES)[1]


def fetch_remote_document(url: str, *, accept: str = "text/html,application/json") -> tuple[str, bytes]:
    """Fetch a page or JSON document under the same guard as an image."""
    return fetch_guarded(url, accept=accept, max_bytes=MAX_DOCUMENT_BYTES)


# ---------------------------------------------------------------------------
# Upload sniffing
# ---------------------------------------------------------------------------

def looks_like_xlsx(head: bytes) -> bool:
    # .xlsx is a zip container; openpyxl will reject anything else anyway, but
    # rejecting it here keeps a malformed file from reaching the parser at all.
    return head.startswith(b"PK\x03\x04")


def looks_like_pdf(head: bytes) -> bool:
    # The %PDF marker is allowed a small leading offset, which real-world
    # writers occasionally produce and every reader tolerates.
    return b"%PDF-" in head[:1024]


def looks_like_text(head: bytes) -> bool:
    if b"\x00" in head:
        return False
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            head.decode(encoding)
        except UnicodeDecodeError:
            continue
        return True
    return False
