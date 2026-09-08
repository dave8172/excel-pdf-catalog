"""The one step in the Shopify pipeline that a model does better than rules.

Everything else about building a catalog from a store URL is deterministic --
see `shopify_catalog`. This is the exception, and it is worth being precise
about why, because "add an LLM" is otherwise a reflex rather than a reason.

Measured on real storefronts, the naive rules pick the wrong logo about half
the time. Death Wish Coffee's homepage has four images tagged "logo" and three
are press badges (BuzzFeed, HuffPost, Yahoo). Tentree's first match is a
Science Based Targets certification mark. Allbirds' favicon is monochrome, so
there is no brand colour to read out of it at all. Each of those is fixable
with one more rule, and the next store breaks the next rule -- which is the
signature of a judgement call, not a parsing problem.

So the model gets exactly that judgement and nothing else: here is a shortlist
of images with their filenames, shapes and palettes; which one is the brand's
own mark, and what are the two colours. It never sees the product data and
never writes a word that reaches the page.

Haiku is the right size for it: the input is a few hundred tokens of already
structured description and the output is a five-field object, so the run costs
a fraction of a cent and adds about a second.

Off by default. `shopify_catalog.read_brand(refine=None)` produces a complete
catalog in a neutral palette, and that is the baseline this has to beat.
"""

from __future__ import annotations

import json
import os

MODEL = "claude-haiku-4-5"

SYSTEM = (
    "You identify a shop's own brand mark among images scraped from its homepage. "
    "Storefronts are full of images that look like logos but belong to someone else: "
    "press mentions, certification marks, payment badges, review-platform logos. "
    "Those are never the answer. Prefer the shop's own wordmark or symbol, in the "
    "widest, most header-like candidate whose filename relates to the shop. If none of "
    "the candidates is plausibly the shop's own mark, return null rather than guessing — "
    "a typeset shop name is a better catalog cover than another company's badge."
)

SCHEMA = {
    "type": "object",
    "properties": {
        "logo_index": {
            "type": ["integer", "null"],
            "description": "Index of the shop's own logo, or null if none of them is one.",
        },
        "primary_color": {
            "type": "string",
            "description": "#RRGGBB for the catalog header band. A strong brand colour, dark "
                           "enough that white text sits on it, or a near-black if the brand is "
                           "monochrome. Take it from the chosen logo's palette when it has one.",
        },
        "secondary_color": {
            "type": "string",
            "description": "#RRGGBB accent for rules and footer text. Must contrast with white paper.",
        },
        "tagline": {
            "type": "string",
            "description": "At most eight words describing what the shop sells, from its title and "
                           "description. Empty string if there is nothing to go on.",
        },
        "reasoning": {
            "type": "string",
            "description": "One short sentence on why that candidate and not the others.",
        },
    },
    "required": ["logo_index", "primary_color", "secondary_color", "tagline", "reasoning"],
    "additionalProperties": False,
}


def available() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY"))


def refine(evidence: dict) -> dict:
    """Choose the logo and the colours. Raises if the model is unreachable.

    `shopify_catalog.read_brand` catches anything raised here and keeps the
    rule-based result, so a missing key or a bad day for the API costs the
    catalog its palette, not its existence.
    """
    import anthropic

    client = anthropic.Anthropic()
    response = client.messages.create(
        model=MODEL,
        max_tokens=500,
        system=SYSTEM,
        messages=[{"role": "user", "content": json.dumps(evidence, ensure_ascii=False)}],
        output_config={"format": {"type": "json_schema", "schema": SCHEMA}},
    )
    text = "".join(block.text for block in response.content if block.type == "text")
    verdict = json.loads(text)
    verdict["_usage"] = (response.usage.input_tokens, response.usage.output_tokens)
    return verdict
