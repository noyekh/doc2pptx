"""
Core data models for the doc2pptx project.

This module defines the base Pydantic models used throughout the application
for data validation and serialization.
"""

from enum import Enum
from pathlib import Path
from typing import Dict, List, Optional, Union
from uuid import uuid4

from pydantic import BaseModel, Field, HttpUrl, model_validator

import logging

logger = logging.getLogger(__name__)


class ContentType(str, Enum):
    """Types of content that can be included in a slide."""

    TEXT = "text"
    BULLET_POINTS = "bullet_points"
    TABLE = "table"
    IMAGE = "image"
    CHART = "chart"
    MERMAID = "mermaid"
    CODE = "code"

    # case-insensitive parsing
    @classmethod
    def _missing_(cls, value):
        if isinstance(value, str):
            value_lower = value.lower()
            for member in cls:
                if member.value == value_lower:
                    return member
        return None


# class SectionType(str, Enum):
#     """Types of sections that can be included in a presentation."""

#     TITLE = "title"
#     INTRODUCTION = "introduction"
#     CONTENT = "content"
#     CONCLUSION = "conclusion"
#     APPENDIX = "appendix"
#     CUSTOM = "custom"
#     AGENDA = "agenda"  # Added agenda type to support the sample JSON file
#     # Types supplémentaires pour les tests
#     SECTION_HEADER = "section_header"
#     BULLET_LIST = "bullet_list"
#     CHART = "chart"
#     TEXT_BLOCKS = "text_blocks"
#     IMAGE_RIGHT = "image_right"
#     TWO_COLUMN = "two_column"
#     TABLE = "table"
#     IMAGE_LEFT = "image_left"
#     HEAT_MAP = "heat_map"
#     QUOTE = "quote"
#     NUMBERED_LIST = "numbered_list"
#     THANK_YOU = "thank_you"
    
#     # case-insensitive parsing
#     @classmethod
#     def _missing_(cls, value):
#         if isinstance(value, str):
#             value_lower = value.lower()
#             for member in cls:
#                 if member.value == value_lower:
#                     return member
#         return None

class SectionType(str, Enum):
    """Types of sections that can be included in a presentation."""

    TITLE = "title"
    INTRODUCTION = "introduction"
    CONTENT = "content"
    CONCLUSION = "conclusion"
    APPENDIX = "appendix"
    CUSTOM = "custom"
    AGENDA = "agenda"  # Added agenda type to support the sample JSON file
    # Types supplémentaires pour les tests
    SECTION_HEADER = "section_header"
    BULLET_LIST = "bullet_list"
    CHART = "chart"
    TEXT_BLOCKS = "text_blocks"
    IMAGE_RIGHT = "image_right"
    TWO_COLUMN = "two_column"
    TABLE = "table"
    IMAGE_LEFT = "image_left"
    HEAT_MAP = "heat_map"
    QUOTE = "quote"
    NUMBERED_LIST = "numbered_list"
    THANK_YOU = "thank_you"
    CODE = "code"  # Added code type
    MERMAID = "mermaid"  # Added mermaid type
    
    # case-insensitive parsing
    @classmethod
    def _missing_(cls, value):
        """
        Handle missing enum values, converting unknown types to CUSTOM.
        
        Args:
            value: The value that doesn't match any enum member.
            
        Returns:
            SectionType: The matched enum member or CUSTOM.
        """
        if isinstance(value, str):
            value_lower = value.lower()
            for member in cls:
                if member.value == value_lower:
                    return member
                
            # Log warning about custom type and return CUSTOM
            logger.warning(f"Unknown section type '{value}' - treating as 'custom'")
            return cls.CUSTOM
        elif value is None:
            logger.warning("None provided as section type - treating as 'custom'")
            return cls.CUSTOM
        else:
            logger.warning(f"Invalid section type {value} of type {type(value)} - treating as 'custom'")
            return cls.CUSTOM

class ImageSource(BaseModel):
    """Source information for an image in a slide."""

    url: Optional[HttpUrl] = None
    path: Optional[Path] = None
    alt_text: Optional[str] = Field(default=None, description="Alternative text for the image")
    query: Optional[str] = Field(
        default=None, description="Query to search for an image on Unsplash"
    )

    @model_validator(mode="after")
    def validate_image_source(self) -> "ImageSource":
        """Ensure that at least one source (url, path, or query) is provided."""
        if not any([self.url, self.path, self.query]):
            raise ValueError("At least one of url, path or query must be provided")
        return self


class TableData(BaseModel):
    """Data for a table in a slide."""

    headers: List[str] = Field(..., description="Headers for the table")
    rows: List[List[str]] = Field(..., description="Rows of data for the table")
    # Ajout d'un champ optionnel pour la compatibilité avec les optimisations IA
    row_count: Optional[int] = Field(None, description="Number of rows in the table")

    @model_validator(mode="after")
    def validate_table_data(self) -> "TableData":
        """
        Ensure that every row has the right number of columns and initialize row_count.
        """
        if not self.headers:
            raise ValueError("Table must have at least one header")

        # Surcharge : on ignore le header 'style:…' à la fin s'il existe
        effective_header_len = len(self.headers)
        if isinstance(self.headers[-1], str) and self.headers[-1].startswith("style:"):
            effective_header_len -= 1

        if effective_header_len == 0:
            raise ValueError("Table must have at least one data column")

        # Validation des lignes
        if self.rows:
            for idx, row in enumerate(self.rows, start=1):
                if len(row) != effective_header_len:
                    raise ValueError(
                        f"Row {idx} has {len(row)} cells, "
                        f"expected {effective_header_len} to match headers"
                    )
            # Initialiser row_count si pas déjà défini
            if self.row_count is None:
                self.row_count = len(self.rows)
        elif self.row_count is None:
            # Si ni rows ni row_count n'est défini, définir row_count à 0
            self.row_count = 0

        return self


class ChartData(BaseModel):
    """Data for a chart in a slide."""

    chart_type: str = Field(..., description="Type of chart (bar, line, pie, etc.)")
    categories: List[str] = Field(..., description="Categories/labels for the chart")
    series: List[Dict[str, Union[str, List[float]]]] = Field(
        ..., description="Series of data for the chart"
    )
    title: Optional[str] = Field(default=None, description="Title of the chart")


class CodeBlock(BaseModel):
    """A block of code to be displayed in a slide."""

    code: str = Field(..., description="The code content")
    language: str = Field(..., description="Programming language of the code")
    line_numbers: bool = Field(default=True, description="Whether to show line numbers")


class MermaidDiagram(BaseModel):
    """A Mermaid diagram to be displayed in a slide."""

    code: str = Field(..., description="The Mermaid diagram code")
    caption: Optional[str] = Field(default=None, description="Caption for the diagram")


class SlideContent(BaseModel):
    """Content to be placed in a specific area of a slide."""

    content_type: ContentType = Field(..., description="Type of content")
    text: Optional[str] = Field(default=None, description="Text content")
    bullet_points: Optional[List[str]] = Field(
        default=None, description="List of bullet points"
    )
    # True ⇒ rend sous forme de puces ; False ⇒ simples paragraphes justifiés
    # Alias « as_bullet » pour matcher l’input existant.
    as_bullets: bool = Field(
        default=True,
        alias="as_bullet",
        description="Render bullet_points as bulleted list"
    )
    
    table: Optional[TableData] = Field(default=None, description="Table data")
    image: Optional[ImageSource] = Field(default=None, description="Image data")
    chart: Optional[ChartData] = Field(default=None, description="Chart data")
    code: Optional[CodeBlock] = Field(default=None, description="Code block")
    mermaid: Optional[MermaidDiagram] = Field(default=None, description="Mermaid diagram")

    @model_validator(mode="after")
    def validate_content_matches_type(self) -> "SlideContent":
        """Ensure that the content field matches the content type."""
        content_field_map = {
            ContentType.TEXT: "text",
            ContentType.BULLET_POINTS: "bullet_points",
            ContentType.TABLE: "table",
            ContentType.IMAGE: "image",
            ContentType.CHART: "chart",
            ContentType.CODE: "code",
            ContentType.MERMAID: "mermaid",
        }

        expected_field = content_field_map.get(self.content_type)
        if expected_field is None:
            raise ValueError(f"Unknown content type: {self.content_type}")

        if getattr(self, expected_field) is None:
            raise ValueError(
                f"Content type is {self.content_type}, but {expected_field} is not provided"
            )

        return self
    
    def default_as_bullets(self) -> "SlideContent":
        """
        Si 'bullet_points' existe et que l'utilisateur n'a pas précisé
        'as_bullet', on force la valeur à True.
        """
        if self.bullet_points is not None and self.as_bullets is None:
            self.as_bullets = True
        return self


class SlideBlock(BaseModel):
    """A block within a slide that contains content and layout information."""

    id: str = Field(default_factory=lambda: str(uuid4()), description="Unique identifier for the block")
    title: Optional[str] = Field(default=None, description="Title of the block")
    content: SlideContent = Field(..., description="Content of the block")
    position: Optional[Dict[str, float]] = Field(
        default=None, description="Position coordinates (left, top, width, height)"
    )
    style: Optional[Dict[str, str]] = Field(
        default=None, description="Style information (colors, fonts, etc.)"
    )


class Slide(BaseModel):
    """A single slide in a presentation."""

    id: str = Field(default_factory=lambda: str(uuid4()), description="Unique identifier for the slide")
    title: str = Field(..., description="Title of the slide")
    layout_name: Optional[str] = Field(default=None, description="Name of the slide layout to use")
    blocks: List[SlideBlock] = Field(..., description="Content blocks in the slide")
    notes: Optional[str] = Field(default=None, description="Speaker notes for the slide")
    background: Optional[Dict[str, str]] = Field(
        default=None, description="Background information (color, image, etc.)"
    )


class Section(BaseModel):
    """A section of a presentation containing multiple slides."""

    id: str = Field(default_factory=lambda: str(uuid4()), description="Unique identifier for the section")
    title: str = Field(..., description="Title of the section")
    type: SectionType = Field(..., description="Type of section")
    slides: List[Slide] = Field(..., description="Slides in the section")
    description: Optional[str] = Field(
        default=None, description="Description of the section content"
    )


class Presentation(BaseModel):
    """Complete presentation model containing multiple sections."""

    id: str = Field(default_factory=lambda: str(uuid4()), description="Unique identifier for the presentation")
    title: str = Field(..., description="Title of the presentation")
    author: Optional[str] = Field(default=None, description="Author of the presentation")
    description: Optional[str] = Field(
        default=None, description="Description of the presentation"
    )
    template_path: Optional[Path] = Field(
        default=None, description="Path to the template PowerPoint file"
    )
    sections: List[Section] = Field(..., description="Sections in the presentation")
    metadata: Optional[Dict[str, str]] = Field(
        default=None, description="Additional metadata for the presentation"
    )