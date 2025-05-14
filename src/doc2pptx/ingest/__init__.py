
# src/doc2pptx/ingest/__init__.py
"""
Data ingestion components for doc2pptx.

This package handles loading and parsing data from various sources like
JSON and Markdown.
"""
from doc2pptx.ingest.json_loader import load_sections, load_presentation
from doc2pptx.ingest.markdown_loader import load_presentation_from_markdown

__all__ = [
    'load_sections',
    'load_presentation',
    'load_presentation_from_markdown'
]