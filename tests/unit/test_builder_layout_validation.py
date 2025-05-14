"""
Unit tests for layout validation and selection in PPTBuilder.
"""
import pytest
from unittest.mock import MagicMock, patch

from doc2pptx.core.models import ContentType, Section, Slide, SlideBlock, SlideContent, TableData, SectionType
from doc2pptx.ppt.builder_v3 import PPTBuilder, LAYOUT_CAPABILITIES


@pytest.fixture
def builder():
    """Create a PPTBuilder instance."""
    return PPTBuilder()


def test_validate_layout_for_content_unknown_layout(builder):
    """Test validating an unknown layout."""
    # Create a slide with an unknown layout
    slide = Slide(
        id="slide1",
        title="Unknown Layout",
        layout_name="NonExistentLayout",
        blocks=[
            SlideBlock(
                id="block1",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="Some text content"
                )
            )
        ]
    )
    
    # Validate layout (should fallback to default)
    with patch('doc2pptx.ppt.builder_v3.logger.warning') as mock_warning:
        result = builder._validate_layout_for_content(slide)
        
        # Should use the default layout
        assert result == "Titre et texte"
        
        # Should log a warning
        mock_warning.assert_called_once()


def test_validate_layout_for_content_too_many_blocks(builder):
    """Test validating a layout with too many blocks."""
    # Create a slide with too many blocks for its layout
    slide = Slide(
        id="slide1",
        title="Too Many Blocks",
        layout_name="Titre et texte",  # Only supports 1 block
        blocks=[
            SlideBlock(id="block1", content=SlideContent(content_type=ContentType.TEXT, text="Text 1")),
            SlideBlock(id="block2", content=SlideContent(content_type=ContentType.TEXT, text="Text 2")),
            SlideBlock(id="block3", content=SlideContent(content_type=ContentType.TEXT, text="Text 3"))
        ]
    )
    
    # Validate layout
    with patch('doc2pptx.ppt.builder_v3.logger.warning') as mock_warning:
        result = builder._validate_layout_for_content(slide)
        
        # Should change to a layout that supports multiple blocks
        assert result == "Titre et 3 colonnes"
        
        # Should log a warning
        mock_warning.assert_called_once()


def test_validate_layout_for_content_table_in_text_layout(builder):
    """Test validating a layout with a table in a text layout."""
    # Create a slide with a table in a text layout
    slide = Slide(
        id="slide1",
        title="Table in Text Layout",
        layout_name="Titre et texte",  # Text layout, not table
        blocks=[
            SlideBlock(
                id="block1",
                content=SlideContent(
                    content_type=ContentType.TABLE,
                    table=TableData(
                        headers=["Col1", "Col2"],
                        rows=[["A", "B"], ["C", "D"]]
                    )
                )
            )
        ]
    )
    
    # Validate layout
    with patch('doc2pptx.ppt.builder_v3.logger.warning') as mock_warning:
        result = builder._validate_layout_for_content(slide)
        
        # Should change to a table layout
        assert result == "Titre et tableau"
        
        # Should log a warning
        mock_warning.assert_called_once()


def test_validate_layout_for_content_text_in_table_layout(builder):
    """Test validating a layout with text in a table layout."""
    # Create a slide with text in a table layout
    slide = Slide(
        id="slide1",
        title="Text in Table Layout",
        layout_name="Titre et tableau",  # Table layout, not text
        blocks=[
            SlideBlock(
                id="block1",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="Some text content"
                )
            )
        ]
    )
    
    # Validate layout
    with patch('doc2pptx.ppt.builder_v3.logger.warning') as mock_warning:
        result = builder._validate_layout_for_content(slide)
        
        # Should change to a text layout
        assert result == "Titre et texte"
        
        # Should log a warning
        mock_warning.assert_called_once()


def test_validate_layout_for_content_appropriate_layout(builder):
    """Test validating an appropriate layout for content."""
    # Create a slide with appropriate layout
    slide = Slide(
        id="slide1",
        title="Appropriate Layout",
        layout_name="Titre et texte",  # Text layout for text content
        blocks=[
            SlideBlock(
                id="block1",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="Some text content"
                )
            )
        ]
    )
    
    # Validate layout (should not change)
    with patch('doc2pptx.ppt.builder_v3.logger.warning') as mock_warning:
        result = builder._validate_layout_for_content(slide)
        
        # Layout should not change
        assert result == "Titre et texte"
        
        # Should not log a warning
        mock_warning.assert_not_called()


def test_validate_layout_for_content_too_many_columns(builder):
    """Test validating a layout with too many text blocks for columns."""
    # Create a slide with more blocks than columns
    slide = Slide(
        id="slide1",
        title="Too Many Columns",
        layout_name="Titre et 3 colonnes",  # Supports 3 blocks
        blocks=[
            SlideBlock(id="block1", content=SlideContent(content_type=ContentType.TEXT, text="Text 1")),
            SlideBlock(id="block2", content=SlideContent(content_type=ContentType.TEXT, text="Text 2")),
            SlideBlock(id="block3", content=SlideContent(content_type=ContentType.TEXT, text="Text 3")),
            SlideBlock(id="block4", content=SlideContent(content_type=ContentType.TEXT, text="Text 4"))
        ]
    )
    
    # Validate layout
    with patch('doc2pptx.ppt.builder_v3.logger.warning') as mock_warning:
        result = builder._validate_layout_for_content(slide)
        
        # Should fallback to text layout
        assert result == "Titre et texte"
        
        # Should log a warning
        mock_warning.assert_called_once()


def test_validate_layout_for_content_no_blocks(builder):
    """Test validating a layout with no content blocks."""
    # Create a slide with no blocks
    slide = Slide(
        id="slide1",
        title="No Blocks",
        layout_name="Titre et texte",
        blocks=[]
    )
    
    # Validate layout (should not change, as zero is less than max_blocks)
    result = builder._validate_layout_for_content(slide)
    
    # Layout should not change
    assert result == "Titre et texte"


def test_needs_section_header(builder):
    """Test determining if a section needs a header slide."""
    # Create different section types
    section_title = Section(
        id="section1",
        title="Title Section",
        type=SectionType.TITLE,
        slides=[
            Slide(id="slide1", title="Slide 1", layout_name="Titre et texte", blocks=[])
        ]
    )
    
    section_content = Section(
        id="section2",
        title="Content Section",
        type=SectionType.CONTENT,
        slides=[
            Slide(id="slide1", title="Slide 1", layout_name="Titre et texte", blocks=[])
        ]
    )
    
    # Test needs_section_header for different section types
    # Currently the implementation returns False for all sections
    assert builder._needs_section_header(section_title) is False
    assert builder._needs_section_header(section_content) is False


def test_create_slide_layout_not_found():
    """Test creating a slide when layout is not found."""
    # Mock PowerPoint presentation
    mock_pptx = MagicMock()
    
    # Setup layouts
    layout1 = MagicMock()
    layout1.name = "Layout1"
    layout2 = MagicMock()
    layout2.name = "Layout2"
    
    mock_pptx.slide_layouts = [layout1, layout2]
    
    builder = PPTBuilder()
    
    # Try to create a slide with non-existent layout
    with patch('doc2pptx.ppt.builder_v3.logger.warning') as mock_warning, \
         patch('doc2pptx.ppt.builder_v3.logger.info') as mock_info:
        slide = builder._create_slide(mock_pptx, "NonExistentLayout")
        
        # Should log a warning
        mock_warning.assert_called_once()
        mock_info.assert_called_once()
        
        # Should use the first available layout
        mock_pptx.slides.add_slide.assert_called_once_with(layout1)


def test_create_slide_layout_found():
    """Test creating a slide when layout is found."""
    # Mock PowerPoint presentation
    mock_pptx = MagicMock()
    
    # Setup layouts
    layout1 = MagicMock()
    layout1.name = "Layout1"
    layout2 = MagicMock()
    layout2.name = "Layout2"
    
    mock_pptx.slide_layouts = [layout1, layout2]
    
    builder = PPTBuilder()
    
    # Create a slide with existing layout
    slide = builder._create_slide(mock_pptx, "Layout2")
    
    # Should use the specified layout
    mock_pptx.slides.add_slide.assert_called_once_with(layout2)


def test_validate_layouts_capabilities():
    """Test that all layouts have proper capabilities defined."""
    # Check that all layouts have the required capabilities
    for layout, capabilities in LAYOUT_CAPABILITIES.items():
        assert "title" in capabilities
        assert "content" in capabilities
        assert "table" in capabilities
        assert "image" in capabilities
        assert "chart" in capabilities
        assert "max_blocks" in capabilities
        assert "description" in capabilities