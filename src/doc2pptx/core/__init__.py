# src/doc2pptx/core/__init__.py
"""
Core components for doc2pptx.

This package contains the core models and settings used throughout the application.
"""

from doc2pptx.core.models import (
    ChartData,
    CodeBlock,
    ContentType,
    ImageSource,
    MermaidDiagram,
    Presentation,
    Section,
    SectionType,
    Slide,
    SlideBlock,
    SlideContent,
    TableData,
)
from doc2pptx.core.settings import Settings, settings

__all__ = [
    "ChartData",
    "CodeBlock",
    "ContentType",
    "ImageSource",
    "MermaidDiagram",
    "Presentation",
    "Section",
    "SectionType",
    "Settings",
    "Slide",
    "SlideBlock",
    "SlideContent",
    "TableData",
    "settings",
]
