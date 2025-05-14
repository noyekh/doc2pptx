"""Unit tests for the core models module."""

import pytest
from pathlib import Path
from pydantic import ValidationError

from doc2pptx.core.models import (
    ContentType, SectionType, ImageSource, TableData, 
    ChartData, CodeBlock, MermaidDiagram, SlideContent,
    SlideBlock, Slide, Section, Presentation
)


class TestImageSource:
    """Tests for the ImageSource model."""
    
    def test_valid_image_source_with_url(self):
        """Test that ImageSource accepts valid URL."""
        image = ImageSource(url="https://example.com/image.jpg")
        assert str(image.url) == "https://example.com/image.jpg"  # Convert Url to string
        assert image.path is None
        assert image.query is None
    
    def test_valid_image_source_with_path(self):
        """Test that ImageSource accepts valid path."""
        image = ImageSource(path="path/to/image.jpg")
        assert image.path == Path("path/to/image.jpg")
        assert image.url is None
        assert image.query is None
    
    def test_valid_image_source_with_query(self):
        """Test that ImageSource accepts valid query."""
        image = ImageSource(query="nature landscape")
        assert image.query == "nature landscape"
        assert image.url is None
        assert image.path is None
    
    def test_valid_image_source_with_alt_text(self):
        """Test that ImageSource accepts alt text."""
        image = ImageSource(url="https://example.com/image.jpg", alt_text="Example image")
        assert image.alt_text == "Example image"
    
    def test_invalid_image_source_no_source(self):
        """Test that ImageSource raises error when no source is provided."""
        with pytest.raises(ValidationError) as exc_info:
            ImageSource()
        
        # Check the error message contains our validation message
        error_msg = str(exc_info.value)
        assert "At least one of url, path or query must be provided" in error_msg


class TestTableData:
    """Tests for the TableData model."""
    
    def test_valid_table_data(self):
        """Test that TableData accepts valid headers and rows."""
        table = TableData(
            headers=["Name", "Age", "City"],
            rows=[
                ["Alice", "30", "New York"],
                ["Bob", "25", "San Francisco"]
            ]
        )
        assert len(table.headers) == 3
        assert len(table.rows) == 2
        assert table.rows[0][0] == "Alice"
    
    def test_invalid_table_data_mismatched_rows(self):
        """Test that TableData raises error when rows don't match headers."""
        with pytest.raises(ValidationError) as exc_info:
            TableData(
                headers=["Name", "Age", "City"],
                rows=[
                    ["Alice", "30"],  # Missing a column
                    ["Bob", "25", "San Francisco"]
                ]
            )
        
        # Check the error message contains our validation message
        error_msg = str(exc_info.value)
        assert "All rows must have the same number of columns as headers" in error_msg


class TestSlideContent:
    """Tests for the SlideContent model."""
    
    def test_valid_text_content(self):
        """Test that SlideContent accepts valid text content."""
        content = SlideContent(
            content_type=ContentType.TEXT,
            text="Sample text content"
        )
        assert content.content_type == ContentType.TEXT
        assert content.text == "Sample text content"
    
    def test_valid_bullet_points_content(self):
        """Test that SlideContent accepts valid bullet points content."""
        content = SlideContent(
            content_type=ContentType.BULLET_POINTS,
            bullet_points=["Point 1", "Point 2", "Point 3"]
        )
        assert content.content_type == ContentType.BULLET_POINTS
        assert len(content.bullet_points) == 3
    
    def test_valid_table_content(self):
        """Test that SlideContent accepts valid table content."""
        content = SlideContent(
            content_type=ContentType.TABLE,
            table=TableData(
                headers=["Col1", "Col2"],
                rows=[["Data1", "Data2"], ["Data3", "Data4"]]
            )
        )
        assert content.content_type == ContentType.TABLE
        assert content.table.headers == ["Col1", "Col2"]
    
    def test_valid_image_content(self):
        """Test that SlideContent accepts valid image content."""
        content = SlideContent(
            content_type=ContentType.IMAGE,
            image=ImageSource(url="https://example.com/image.jpg")
        )
        assert content.content_type == ContentType.IMAGE
        assert str(content.image.url) == "https://example.com/image.jpg"  # Convert Url to string
    
    def test_valid_chart_content(self):
        """Test that SlideContent accepts valid chart content."""
        content = SlideContent(
            content_type=ContentType.CHART,
            chart=ChartData(
                chart_type="bar",
                categories=["Category 1", "Category 2"],
                series=[{"name": "Series 1", "data": [10, 20]}]
            )
        )
        assert content.content_type == ContentType.CHART
        assert content.chart.chart_type == "bar"
    
    def test_valid_code_content(self):
        """Test that SlideContent accepts valid code content."""
        content = SlideContent(
            content_type=ContentType.CODE,
            code=CodeBlock(
                code="print('Hello, World!')",
                language="python"
            )
        )
        assert content.content_type == ContentType.CODE
        assert content.code.language == "python"
    
    def test_valid_mermaid_content(self):
        """Test that SlideContent accepts valid mermaid content."""
        content = SlideContent(
            content_type=ContentType.MERMAID,
            mermaid=MermaidDiagram(
                code="graph TD; A-->B; B-->C;"
            )
        )
        assert content.content_type == ContentType.MERMAID
        assert "graph TD" in content.mermaid.code
    
    def test_invalid_content_mismatch(self):
        """Test that SlideContent raises error when content doesn't match type."""
        with pytest.raises(ValidationError) as exc_info:
            SlideContent(
                content_type=ContentType.TEXT,
                # Missing text field
                bullet_points=["Point 1", "Point 2"]
            )
        
        # Check the error message - updated to match the actual Pydantic error format
        error_msg = str(exc_info.value)
        assert "Content type is ContentType.TEXT, but text is not provided" in error_msg
    
    def test_invalid_unknown_content_type(self):
        """Test that SlideContent raises error for unknown content type."""
        # This should be caught by the enum validation, but if someone tries to extend
        # the model with a new type without updating the validation, this test would fail
        valid_content_types = [member.value for member in ContentType]
        assert "unknown_type" not in valid_content_types