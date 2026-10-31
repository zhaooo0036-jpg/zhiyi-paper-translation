# -*- coding: utf-8 -*-
"""纸译 zhiyi — turn an English paper PDF into a Chinese .docx translation.

Pipeline: extract -> probe -> crop -> build -> verify, with one shared glossary
so that the same term always gets the same Chinese name across documents.
"""

__version__ = "0.1.0"

from .docx import DocBuilder  # noqa: F401,E402
