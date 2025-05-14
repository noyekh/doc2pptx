"""
Unit tests for core models and settings in doc2pptx.

This module contains tests for the models and settings defined in
the core package.
"""

import os
from pathlib import Path
from unittest.mock import patch

import pytest
from pydantic import ValidationError

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
from doc2pptx.core.settings import Settings


# Model tests
class TestImageSource:
    """Tests for the ImageSource model."""

    def test_valid_url(self):
        """Test creating an ImageSource with a valid URL."""
        source = ImageSource(url="https://example.com/image.jpg")
        assert str(source.url) == "https://example.com/image.jpg"
        assert source.path is None
        assert source.query is None

    def test_valid_path(self):
        """Test creating an ImageSource with a valid path."""
        source = ImageSource(path=Path("/path/to/image.jpg"))
        assert source.url is None
        assert source.path == Path("/path/to/image.jpg")
        assert source.query is None

    def test_valid_query(self):
        """Test creating an ImageSource with a valid query."""
        source = ImageSource(query="nature landscape")
        assert source.url is None
        assert source.path is None
        assert source.query == "nature landscape"

    def test_invalid_no_source(self):
        """Test that ValidationError is raised when no source is provided."""
        with pytest.raises(ValueError, match="At least one of url, path or query must be provided"):
            ImageSource()


class TestTableData:
    """Tests for the TableData model."""

    def test_valid_table(self):
        """Test creating a valid TableData."""
        table = TableData(
            headers=["Name", "Age", "City"],
            rows=[
                ["Alice", "30", "New York"],
                ["Bob", "25", "San Francisco"],
            ],
        )
        assert table.headers == ["Name", "Age", "City"]
        assert len(table.rows) == 2
        assert table.rows[0] == ["Alice", "30", "New York"]

    def test_invalid_row_length(self):
        """Test that ValidationError is raised when row length doesn't match headers."""
        with pytest.raises(ValueError, match="All rows must have the same number of columns as headers"):
            TableData(
                headers=["Name", "Age", "City"],
                rows=[
                    ["Alice", "30"],  # Missing city
                    ["Bob", "25", "San Francisco"],
                ],
            )


class TestChartData:
    """Tests for the ChartData model."""

    def test_valid_chart(self):
        """Test creating a valid ChartData."""
        chart = ChartData(
            chart_type="bar",
            categories=["Jan", "Feb", "Mar"],
            series=[
                {"name": "Sales", "data": [10, 15, 20]},
                {"name": "Costs", "data": [5, 8, 12]},
            ],
            title="Quarterly Results",
        )
        assert chart.chart_type == "bar"
        assert chart.categories == ["Jan", "Feb", "Mar"]
        assert len(chart.series) == 2
        assert chart.title == "Quarterly Results"


class TestCodeBlock:
    """Tests for the CodeBlock model."""

    def test_valid_code_block(self):
        """Test creating a valid CodeBlock."""
        code = CodeBlock(
            code="def hello_world():\n    print('Hello, world!')",
            language="python",
        )
        assert "Hello, world!" in code.code
        assert code.language == "python"
        assert code.line_numbers is True

    def test_no_line_numbers(self):
        """Test creating a CodeBlock with line_numbers=False."""
        code = CodeBlock(
            code="function helloWorld() { console.log('Hello, world!'); }",
            language="javascript",
            line_numbers=False,
        )
        assert code.line_numbers is False


class TestMermaidDiagram:
    """Tests for the MermaidDiagram model."""

    def test_valid_mermaid(self):
        """Test creating a valid MermaidDiagram."""
        diagram = MermaidDiagram(
            code="graph TD;\n    A-->B;\n    A-->C;\n    B-->D;\n    C-->D;",
            caption="Simple flowchart",
        )
        assert "graph TD" in diagram.code
        assert diagram.caption == "Simple flowchart"


class TestSlideContent:
    """Tests for the SlideContent model."""

    def test_text_content(self):
        """Test creating a SlideContent with text."""
        content = SlideContent(
            content_type=ContentType.TEXT,
            text="This is a text content.",
        )
        assert content.content_type == ContentType.TEXT
        assert content.text == "This is a text content."

    def test_bullet_points(self):
        """Test creating a SlideContent with bullet points."""
        content = SlideContent(
            content_type=ContentType.BULLET_POINTS,
            bullet_points=["Point 1", "Point 2", "Point 3"],
        )
        assert content.content_type == ContentType.BULLET_POINTS
        assert len(content.bullet_points) == 3

    def test_table_content(self):
        """Test creating a SlideContent with a table."""
        content = SlideContent(
            content_type=ContentType.TABLE,
            table=TableData(
                headers=["Name", "Age"], rows=[["Alice", "30"], ["Bob", "25"]]
            ),
        )
        assert content.content_type == ContentType.TABLE
        assert content.table.headers == ["Name", "Age"]

    def test_image_content(self):
        """Test creating a SlideContent with an image."""
        content = SlideContent(
            content_type=ContentType.IMAGE,
            image=ImageSource(url="https://example.com/image.jpg"),
        )
        assert content.content_type == ContentType.IMAGE
        assert str(content.image.url) == "https://example.com/image.jpg"

    def test_chart_content(self):
        """Test creating a SlideContent with a chart."""
        content = SlideContent(
            content_type=ContentType.CHART,
            chart=ChartData(
                chart_type="line",
                categories=["Jan", "Feb", "Mar"],
                series=[{"name": "Sales", "data": [10, 15, 20]}],
            ),
        )
        assert content.content_type == ContentType.CHART
        assert content.chart.chart_type == "line"

    def test_code_content(self):
        """Test creating a SlideContent with code."""
        content = SlideContent(
            content_type=ContentType.CODE,
            code=CodeBlock(code="print('Hello')", language="python"),
        )
        assert content.content_type == ContentType.CODE
        assert content.code.language == "python"

    def test_mermaid_content(self):
        """Test creating a SlideContent with a Mermaid diagram."""
        content = SlideContent(
            content_type=ContentType.MERMAID,
            mermaid=MermaidDiagram(code="graph TD;\n    A-->B;"),
        )
        assert content.content_type == ContentType.MERMAID
        assert "graph TD" in content.mermaid.code

    def test_content_type_mismatch(self):
        """Test that ValidationError is raised when content doesn't match type."""
        with pytest.raises(Exception, match="Content type is ContentType.TEXT, but text is not provided"):
            SlideContent(
                content_type=ContentType.TEXT,
                bullet_points=["This should be text, not bullet points"],
            )


class TestSlideBlock:
    """Tests for the SlideBlock model."""

    def test_minimal_slide_block(self):
        """Test creating a minimal SlideBlock."""
        block = SlideBlock(
            id="block1",
            content=SlideContent(
                content_type=ContentType.TEXT,
                text="This is a text block.",
            ),
        )
        assert block.id == "block1"
        assert block.title is None
        assert block.content.content_type == ContentType.TEXT
        assert block.position is None
        assert block.style is None

    def test_complete_slide_block(self):
        """Test creating a complete SlideBlock with all fields."""
        block = SlideBlock(
            id="block2",
            title="Sample Block",
            content=SlideContent(
                content_type=ContentType.TEXT,
                text="This is a complete text block.",
            ),
            position={"left": 0.1, "top": 0.2, "width": 0.8, "height": 0.3},
            style={"color": "#000000", "font-size": "18pt"},
        )
        assert block.title == "Sample Block"
        assert block.position["left"] == 0.1
        assert block.style["color"] == "#000000"

class TestSlide:
    """Tests for the Slide model."""

    def test_minimal_slide(self):
        """Test creating a minimal Slide."""
        slide = Slide(
            id="slide1",
            title="Introduction",
            layout_name="Title Slide",
            blocks=[
                SlideBlock(
                    id="block1",
                    content=SlideContent(
                        content_type=ContentType.TEXT,
                        text="Introduction text",
                    ),
                ),
            ],
        )
        assert slide.id == "slide1"
        assert slide.title == "Introduction"
        assert slide.layout_name == "Title Slide"
        assert len(slide.blocks) == 1
        assert slide.notes is None
        assert slide.background is None

    def test_complete_slide(self):
        """Test creating a complete Slide with all fields."""
        slide = Slide(
            id="slide2",
            title="Content Slide",
            layout_name="Two Content",
            blocks=[
                SlideBlock(
                    id="block1",
                    content=SlideContent(
                        content_type=ContentType.TEXT,
                        text="Left content",
                    ),
                ),
                SlideBlock(
                    id="block2",
                    content=SlideContent(
                        content_type=ContentType.IMAGE,
                        image=ImageSource(url="https://example.com/image.jpg"),
                    ),
                ),
            ],
            notes="Speaker notes for this slide",
            background={"color": "#FFFFFF"},
        )
        assert slide.title == "Content Slide"
        assert len(slide.blocks) == 2
        assert slide.notes == "Speaker notes for this slide"
        assert slide.background["color"] == "#FFFFFF"


class TestSection:
    """Tests for the Section model."""

    def test_minimal_section(self):
        """Test creating a minimal Section."""
        section = Section(
            id="section1",
            title="Introduction",
            type=SectionType.INTRODUCTION,
            slides=[
                Slide(
                    id="slide1",
                    title="Welcome",
                    layout_name="Title Slide",
                    blocks=[
                        SlideBlock(
                            id="block1",
                            content=SlideContent(
                                content_type=ContentType.TEXT,
                                text="Welcome text",
                            ),
                        ),
                    ],
                ),
            ],
        )
        assert section.id == "section1"
        assert section.title == "Introduction"
        assert section.type == SectionType.INTRODUCTION
        assert len(section.slides) == 1
        assert section.description is None

    def test_complete_section(self):
        """Test creating a complete Section with all fields."""
        section = Section(
            id="section2",
            title="Main Content",
            type=SectionType.CONTENT,
            slides=[
                Slide(
                    id="slide1",
                    title="Slide 1",
                    layout_name="Titre et texte",
                    blocks=[
                        SlideBlock(
                            id="block1",
                            content=SlideContent(
                                content_type=ContentType.TEXT,
                                text="Content text",
                            ),
                        ),
                    ],
                ),
                Slide(
                    id="slide2",
                    title="Slide 2",
                    layout_name="Two Content",
                    blocks=[
                        SlideBlock(
                            id="block1",
                            content=SlideContent(
                                content_type=ContentType.BULLET_POINTS,
                                bullet_points=["Point 1", "Point 2"],
                            ),
                        ),
                    ],
                ),
            ],
            description="Main content section with key points",
        )
        assert section.title == "Main Content"
        assert section.type == SectionType.CONTENT
        assert len(section.slides) == 2
        assert section.description == "Main content section with key points"


class TestPresentation:
    """Tests for the Presentation model."""

    def test_minimal_presentation(self):
        """Test creating a minimal Presentation."""
        presentation = Presentation(
            title="Test Presentation",
            sections=[
                Section(
                    id="section1",
                    title="Introduction",
                    type=SectionType.INTRODUCTION,
                    slides=[
                        Slide(
                            id="slide1",
                            title="Welcome",
                            layout_name="Title Slide",
                            blocks=[
                                SlideBlock(
                                    id="block1",
                                    content=SlideContent(
                                        content_type=ContentType.TEXT,
                                        text="Welcome text",
                                    ),
                                ),
                            ],
                        ),
                    ],
                ),
            ],
        )
        assert presentation.title == "Test Presentation"
        assert presentation.author is None
        assert presentation.description is None
        assert presentation.template_path is None
        assert len(presentation.sections) == 1
        assert presentation.metadata is None

    def test_complete_presentation(self):
        """Test creating a complete Presentation with all fields."""
        presentation = Presentation(
            title="Complete Presentation",
            author="Test Author",
            description="Test presentation with all fields",
            template_path=Path("/path/to/template.pptx"),
            sections=[
                Section(
                    id="section1",
                    title="Introduction",
                    type=SectionType.INTRODUCTION,
                    slides=[
                        Slide(
                            id="slide1",
                            title="Welcome",
                            layout_name="Title Slide",
                            blocks=[
                                SlideBlock(
                                    id="block1",
                                    content=SlideContent(
                                        content_type=ContentType.TEXT,
                                        text="Welcome text",
                                    ),
                                ),
                            ],
                        ),
                    ],
                ),
                Section(
                    id="section2",
                    title="Content",
                    type=SectionType.CONTENT,
                    slides=[
                        Slide(
                            id="slide2",
                            title="Content Slide",
                            layout_name="Titre et texte",
                            blocks=[
                                SlideBlock(
                                    id="block1",
                                    content=SlideContent(
                                        content_type=ContentType.BULLET_POINTS,
                                        bullet_points=["Point 1", "Point 2"],
                                    ),
                                ),
                            ],
                        ),
                    ],
                ),
            ],
            metadata={"created_at": "2025-04-29", "version": "1.0"},
        )
        assert presentation.title == "Complete Presentation"
        assert presentation.author == "Test Author"
        assert presentation.description == "Test presentation with all fields"
        assert presentation.template_path == Path("/path/to/template.pptx")
        assert len(presentation.sections) == 2
        assert presentation.metadata["version"] == "1.0"


# Settings tests
class TestSettings:
    """Tests for the Settings class."""

    @patch.dict(os.environ, {
        "OPENAI_API_KEY": "test-key-openai",
        "UNSPLASH_ACCESS_KEY": "test-access-key",
        "UNSPLASH_SECRET_KEY": "test-secret-key",
        "TEMPLATES_DIR": "/test/templates",
        "OUTPUT_DIR": "/test/output",
        "CACHE_DIR": "/test/cache",
        "MERMAID_CLI_PATH": "/usr/bin/mmdc",
        "OPENAI_MODEL": "gpt-4o",
        "OPENAI_TEMPERATURE": "0.5",
        "DEBUG": "True",
        "LAYOUT_RULES_PATH": "/test/layout/rules.yaml",
        "DOC2PPTX_CUSTOM_SETTING": "test-value"
    })
    def test_load_from_env(self):
        """Test loading settings from environment variables."""
        with patch("pathlib.Path.exists", return_value=True), \
             patch("pathlib.Path.is_dir", return_value=True), \
             patch("pathlib.Path.mkdir"), \
             patch("pathlib.Path.is_absolute", return_value=True):  # Pour éviter la conversion en chemin absolu
            settings = Settings()
            settings.load_custom_env_vars()

            assert settings.openai_api_key == "test-key-openai"
            assert settings.unsplash_access_key == "test-access-key"
            assert settings.unsplash_secret_key == "test-secret-key"
            assert str(settings.templates_dir).replace("C:", "").replace("\\", "/") == "/test/templates"
            assert str(settings.output_dir).replace("C:", "").replace("\\", "/") == "/test/output"
            assert str(settings.cache_dir).replace("C:", "").replace("\\", "/") == "/test/cache"
            assert settings.mermaid_cli_path == "/usr/bin/mmdc"
            assert settings.openai_model == "gpt-4o"
            assert settings.openai_temperature == 0.5
            assert settings.debug is True
            assert str(settings.layout_rules_path).replace("C:", "").replace("\\", "/") == "/test/layout/rules.yaml"
            assert settings.custom_env_vars["custom_setting"] == "test-value"

    @patch.dict(os.environ, {"OPENAI_API_KEY": "test-key-openai"})
    def test_default_values(self):
        """Test default values when environment variables are not set."""
        with patch("pathlib.Path.exists", return_value=False), \
             patch("pathlib.Path.mkdir"):
            settings = Settings()

            assert settings.openai_api_key == "test-key-openai"
            assert settings.unsplash_access_key is None
            assert settings.unsplash_secret_key is None
            assert settings.templates_dir == Path.cwd() / "templates"
            assert settings.output_dir == Path.cwd() / "output"
            assert settings.cache_dir == Path.cwd() / "cache"
            assert settings.mermaid_cli_path == "mmdc"
            assert settings.openai_model == "gpt-4o"
            assert settings.openai_temperature == 0.0
            assert settings.debug is False
            assert settings.layout_rules_path == Path.cwd() / "layout/rules.yaml"
            assert settings.custom_env_vars == {}

    @patch.dict(os.environ, {"OPENAI_API_KEY": "test-key-openai"})
    def test_create_directories(self):
        """Test that directories are created if they don't exist."""
        with patch("pathlib.Path.exists", return_value=False), \
             patch("pathlib.Path.mkdir") as mock_mkdir:
            settings = Settings()
            
            # Check that mkdir was called for each directory path
            assert mock_mkdir.call_count >= 4  # templates_dir, output_dir, cache_dir, layout_rules_path parent

    @patch.dict(os.environ, {"OPENAI_API_KEY": "test-key-openai", "TEMPLATES_DIR": "/existing/file.txt"})
    def test_path_is_file(self):
        """Test that ValueError is raised when a path exists but is a file."""
        with patch("pathlib.Path.exists", return_value=True), \
             patch("pathlib.Path.is_dir", return_value=False):
            with pytest.raises(ValueError, match="exists but is not a directory"):
                Settings()

    @patch.dict(os.environ, {}, clear=True)  # Clear=True vide complètement les variables d'environnement
    def test_missing_required_keys(self):
        """Test that ValidationError is raised when required keys are missing."""
        from pydantic import ValidationError
        
        with pytest.raises(ValidationError):
            Settings()