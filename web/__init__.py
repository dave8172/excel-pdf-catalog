"""Web-facing concerns for the hosted exporter: guards, quotas, templates.

`catalog_exporter.py` is the engine and knows nothing about HTTP. Everything in
this package exists because the engine is now fed by strangers rather than by
its author.
"""
