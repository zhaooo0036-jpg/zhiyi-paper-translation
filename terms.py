# -*- coding: utf-8 -*-
"""Glossary store: one shared, editable JSON table for every translated document."""
import io, json, os

_HERE = os.path.dirname(os.path.abspath(__file__))
_DEFAULT = os.path.join(_HERE, "glossary.json")


def load(path=None):
    with io.open(path or _DEFAULT, encoding="utf-8") as fh:
        raw = json.load(fh)
    out = {}
    for k, v in raw.items():
        out[k] = (v["zh"], v["en"], v.get("note", ""), v.get("abbr"))
    return out


TERMS = load()


def override(path):
    """Merge an extra glossary file into the active table (project-specific terms)."""
    TERMS.update(load(path))
    return TERMS
