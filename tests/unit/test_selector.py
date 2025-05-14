"""
Unit tests for the layout selector module.
"""
import os
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
import yaml
from pptx import Presentation as PptxPresentation

from doc2pptx.core.models import Section, Slide, SlideBlock, SlideContent, SectionType, ContentType
from doc2pptx.layout.selector import LayoutSelector


@pytest.fixture
def temp_rules_file():
    """Create a temporary rules file for testing."""
    rules_content = {
        "default_layout": "Titre et texte",
        "section_types": {
            "title": "Diapositive de titre",
            "introduction": "Introduction",
            "content": "Titre et texte",
            "conclusion": "Chapitre 1",
            "appendix": "Titre et texte",
            "custom": "Titre et texte",
            "agenda": "Titre et texte"
        },
        "content_types": {
            "text": "Titre et texte",
            "bullet_points": "Titre et texte",
            "table": "Titre et tableau",
            "image": "Titre et texte 1 visuel gauche",
            "chart": "Titre et texte 1 histogramme",
            "mermaid": "Titre et texte 1 histogramme",
            "code": "Titre et texte"
        },
        "content_patterns": {
            "^thank you": "Chapitre 1",
            "agenda": "Titre et texte",
            "two columns": "Titre et 3 colonnes"
        },
        "content_combinations": [
            {
                "requires": {
                    "content_types": ["image", "text"],
                    "block_count": 2
                },
                "layout": "Titre et texte 1 visuel gauche"
            },
            {
                "requires": {
                    "content_types": ["bullet_points"],
                    "title_pattern": "^key points"
                },
                "layout": "Titre et texte"
            }
        ],
        "multi_block_layout": "Titre et texte",
        "two_block_layout": "Titre et 3 colonnes"
    }
    
    with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as temp_file:
        yaml.dump(rules_content, temp_file)
        temp_path = temp_file.name
    
    yield Path(temp_path)
    
    # Clean up
    os.unlink(temp_path)


@pytest.fixture
def mock_template():
    """Create a mock PowerPoint template with layout names."""
    template = MagicMock(spec=PptxPresentation)
    
    # Create mock layouts
    layout_names = [
        "Diapositive de titre", 
        "Titre et texte",
        "Chapitre 1",
        "Titre et tableau",
        "Titre et texte 1 visuel gauche",
        "Titre et texte 1 histogramme",
        "Titre et 3 colonnes",
        "Titre et texte",
        "Introduction"
    ]
    
    mock_layouts = []
    for name in layout_names:
        mock_layout = MagicMock()
        mock_layout.name = name
        mock_layouts.append(mock_layout)
    
    # Create a slide_layouts property
    template.slide_layouts = mock_layouts
    
    return template


@pytest.fixture
def layout_selector(temp_rules_file, mock_template):
    """Create a LayoutSelector instance with test rules and template."""
    return LayoutSelector(rules_path=temp_rules_file, template=mock_template)


@pytest.fixture
def create_section():
    """Create a test section with the specified type and slides."""
    def _create_section(section_type, slides=None):
        section_slides = []
        
        if slides:
            for slide_data in slides:
                blocks = []
                for block_data in slide_data.get("blocks", []):
                    content_type = block_data.get("content_type", "text")
                    content_dict = {
                        "content_type": ContentType(content_type)
                    }
                    
                    # Add content based on type
                    if content_type == "text":
                        content_dict["text"] = block_data.get("text", "")
                    elif content_type == "bullet_points":
                        content_dict["bullet_points"] = block_data.get("bullet_points", [])
                    elif content_type == "image":
                        content_dict["image"] = block_data.get("image", {"query": "test", "alt_text": "test"})
                    
                    blocks.append(
                        SlideBlock(
                            id=block_data.get("id", "block-id"),
                            title=block_data.get("title"),
                            content=SlideContent(**content_dict)
                        )
                    )
                
                section_slides.append(
                    Slide(
                        id=slide_data.get("id", "slide-id"),
                        title=slide_data.get("title", "Test Slide"),
                        layout_name=slide_data.get("layout_name", "auto"),
                        blocks=blocks
                    )
                )
        
        return Section(
            id="section-id",
            title="Test Section",
            type=SectionType(section_type),
            slides=section_slides if section_slides else []
        )
    
    return _create_section


def test_layout_selector_init_default_path():
    """Test LayoutSelector initialization with default path."""
    with patch("pathlib.Path.exists", return_value=True), \
         patch("builtins.open", MagicMock()), \
         patch("yaml.safe_load", return_value={"section_types": {}, "content_types": {}}):
        selector = LayoutSelector()
        assert isinstance(selector, LayoutSelector)
        assert selector.rules == {"section_types": {}, "content_types": {}}


def test_layout_selector_init_file_not_found():
    """Test LayoutSelector initialization with non-existent file."""
    with pytest.raises(FileNotFoundError):
        LayoutSelector(rules_path="non_existent_file.yaml")


def test_layout_selector_init_invalid_yaml(temp_rules_file):
    """Test LayoutSelector initialization with invalid YAML."""
    # Write invalid YAML to the temp file
    with open(temp_rules_file, "w") as f:
        f.write("invalid: yaml: content:\n  - missing colon\n")
    
    with pytest.raises(ValueError):
        LayoutSelector(rules_path=temp_rules_file)


def test_layout_selector_validate_rules(temp_rules_file):
    """Test validation of rules against template."""
    # Write rules with layouts that don't exist in the template
    rules_content = {
        "section_types": {"title": "Non-existent Layout"},
        "content_types": {"text": "Another Non-existent Layout"}
    }
    
    with open(temp_rules_file, "w") as f:
        yaml.dump(rules_content, f)
    
    mock_template = MagicMock(spec=PptxPresentation)
    mock_template.slide_layouts = [MagicMock(name="Title Slide")]
    
    with pytest.raises(ValueError):
        LayoutSelector(rules_path=temp_rules_file, template=mock_template)


def test_get_layout_name_with_existing_layout(layout_selector, create_section):
    """Test get_layout_name with a slide that already has a valid layout."""
    section = create_section("content")
    slide = Slide(
        id="slide-id",
        title="Test Slide",
        layout_name="Titre et texte",  # This layout exists in mock_template
        blocks=[
            SlideBlock(
                id="block-id",
                content=SlideContent(content_type=ContentType.TEXT, text="Test content")
            )
        ]
    )
    
    layout_name = layout_selector.get_layout_name(section, slide)
    assert layout_name == "Titre et texte"


def test_get_layout_name_with_invalid_layout(layout_selector, create_section):
    """Test get_layout_name with a slide that has an invalid layout."""
    section = create_section("content")
    slide = Slide(
        id="slide-id",
        title="Test Slide",
        layout_name="Non-existent Layout",  # This layout doesn't exist
        blocks=[
            SlideBlock(
                id="block-id",
                content=SlideContent(content_type=ContentType.TEXT, text="Test content")
            )
        ]
    )
    
    # Should fall back to content-based selection
    layout_name = layout_selector.get_layout_name(section, slide)
    assert layout_name == "Titre et texte"  # Based on TEXT content type


def test_get_layout_name_section_only(layout_selector, create_section):
    """Test get_layout_name with section only."""
    # For different section types
    section_types = {
        "title": "Diapositive de titre",
        "introduction": "Introduction",
        "content": "Titre et texte",
        "conclusion": "Chapitre 1",
        "appendix": "Titre et texte",
        "custom": "Titre et texte",
        "agenda": "Titre et texte"
    }
    
    for section_type, expected_layout in section_types.items():
        section = create_section(section_type)
        layout_name = layout_selector.get_layout_name(section)
        assert layout_name == expected_layout


def test_get_layout_name_with_existing_layout(layout_selector, create_section):
    """Test get_layout_name with a slide that already has a valid layout."""
    section = create_section("content")
    slide = Slide(
        id="slide-id",
        title="Test Slide",
        layout_name="Titre et texte",  # Changer pour un layout qui existe dans nos règles
        blocks=[
            SlideBlock(
                id="block-id",
                content=SlideContent(content_type=ContentType.TEXT, text="Test content")
            )
        ]
    )
    
    layout_name = layout_selector.get_layout_name(section, slide)
    assert layout_name == "Titre et texte"  # S'attendre au même layout


# Correction 2: Modifier test_get_layout_name_with_invalid_layout
def test_get_layout_name_with_invalid_layout(layout_selector, create_section):
    """Test get_layout_name with a slide that has an invalid layout."""
    section = create_section("content")
    slide = Slide(
        id="slide-id",
        title="Test Slide",
        layout_name="Non-existent Layout",  # Ce layout n'existe pas
        blocks=[
            SlideBlock(
                id="block-id",
                content=SlideContent(content_type=ContentType.TEXT, text="Test content")
            )
        ]
    )
    
    # Should fall back to content-based selection
    layout_name = layout_selector.get_layout_name(section, slide)
    assert layout_name == "Titre et texte"  # S'attendre au layout pour le type TEXT


# Correction 3: Modifier test_get_layout_from_content_type
def test_get_layout_from_content_type(layout_selector, create_section):
    """Test layout selection based on content type."""
    section = create_section("content")
    
    # Test different content types
    content_types = {
        "text": "Titre et texte",
        "bullet_points": "Titre et texte",
        "table": "Titre et tableau",
        "image": "Titre et texte 1 visuel gauche",
        "chart": "Titre et texte 1 histogramme",
        "mermaid": "Titre et texte 1 histogramme",
        "code": "Titre et texte"
    }
    
    for content_type, expected_layout in content_types.items():
        # Créer le contenu approprié en fonction du type
        content_kwargs = {}
        if content_type == "text":
            content_kwargs = {"text": "Test content"}
        elif content_type == "bullet_points":
            content_kwargs = {"bullet_points": ["Item 1", "Item 2"]}
        elif content_type == "table":
            content_kwargs = {"table": {"headers": ["Col1", "Col2"], "rows": [["Val1", "Val2"]]}}
        elif content_type == "image":
            content_kwargs = {"image": {"query": "test", "alt_text": "test"}}
        elif content_type == "chart":
            content_kwargs = {"chart": {"chart_type": "bar", "categories": ["A", "B"], "series": [{"name": "Series1", "data": [1.0, 2.0]}]}}
        elif content_type == "mermaid":
            content_kwargs = {"mermaid": {"code": "graph TD; A-->B;"}}
        elif content_type == "code":
            content_kwargs = {"code": {"code": "print('hello')", "language": "python"}}
        
        slide = Slide(
            id="slide-id",
            title="Test Slide",
            layout_name="auto",
            blocks=[
                SlideBlock(
                    id="block-id",
                    content=SlideContent(
                        content_type=ContentType(content_type),
                        **content_kwargs
                    )
                )
            ]
        )
        
        layout_name = layout_selector.get_layout_name(section, slide)
        assert layout_name == expected_layout, f"Failed for content type {content_type}"


def test_get_layout_from_content_pattern(layout_selector, create_section):
    """Test layout selection based on content pattern."""
    section = create_section("content")
    
    # Test different content patterns
    patterns = {
        "Thank you for your attention!": "Chapitre 1",
        "Here's our agenda for today": "Titre et texte",
        "Let's compare these two columns": "Titre et 3 colonnes"
    }
    
    for content_text, expected_layout in patterns.items():
        slide = Slide(
            id="slide-id",
            title="Test Slide",
            layout_name="auto",
            blocks=[
                SlideBlock(
                    id="block-id",
                    content=SlideContent(
                        content_type=ContentType.TEXT,
                        text=content_text
                    )
                )
            ]
        )
        
        layout_name = layout_selector.get_layout_name(section, slide)
        assert layout_name == expected_layout, f"Failed for pattern in: {content_text}"


def test_get_layout_from_content_combination(layout_selector, create_section):
    """Test layout selection based on content combinations."""
    section = create_section("content")
    
    # Test image + text combination
    slide = Slide(
        id="slide-id",
        title="Test Slide",
        layout_name="auto",
        blocks=[
            SlideBlock(
                id="block1-id",
                content=SlideContent(
                    content_type=ContentType.IMAGE,
                    image={"query": "test", "alt_text": "test"}
                )
            ),
            SlideBlock(
                id="block2-id",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="Caption for the image"
                )
            )
        ]
    )
    
    layout_name = layout_selector.get_layout_name(section, slide)
    assert layout_name == "Titre et texte 1 visuel gauche"
    
    # Test title pattern requirement
    slide = Slide(
        id="slide-id",
        title="Key Points to Remember",
        layout_name="auto",
        blocks=[
            SlideBlock(
                id="block-id",
                content=SlideContent(
                    content_type=ContentType.BULLET_POINTS,
                    bullet_points=["Point 1", "Point 2", "Point 3"]
                )
            )
        ]
    )
    
    layout_name = layout_selector.get_layout_name(section, slide)
    assert layout_name == "Titre et texte"


def test_get_layout_based_on_block_count(layout_selector, create_section):
    """Test layout selection based on number of blocks."""
    section = create_section("content")
    
    # Test with two blocks
    slide = Slide(
        id="slide-id",
        title="Test Slide",
        layout_name="auto",
        blocks=[
            SlideBlock(
                id="block1-id",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="First block"
                )
            ),
            SlideBlock(
                id="block2-id",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="Second block"
                )
            )
        ]
    )
    
    layout_name = layout_selector.get_layout_name(section, slide)
    assert layout_name == "Titre et 3 colonnes"
    
    # Test with more than three blocks
    slide = Slide(
        id="slide-id",
        title="Test Slide",
        layout_name="auto",
        blocks=[
            SlideBlock(
                id="block1-id",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="First block"
                )
            ),
            SlideBlock(
                id="block2-id",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="Second block"
                )
            ),
            SlideBlock(
                id="block3-id",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="Third block"
                )
            ),
            SlideBlock(
                id="block4-id",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="Fourth block"
                )
            )
        ]
    )
    
    layout_name = layout_selector.get_layout_name(section, slide)
    assert layout_name == "Titre et texte"


def test_default_layout_fallback(layout_selector, create_section):
    """Test fallback to default layout when no rules match."""
    # Create a section with an unknown type that's not in the rules
    section = create_section("custom")
    
    # Create a slide with an unknown content type
    slide = Slide(
        id="slide-id",
        title="Test Slide",
        layout_name="auto",
        blocks=[]  # No blocks, so can't match on content type
    )
    
    layout_name = layout_selector.get_layout_name(section, slide)
    assert layout_name == "Titre et texte"  # Default layout