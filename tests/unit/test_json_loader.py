"""
Unit tests for the json_loader module in doc2pptx.
This module tests all functionality for loading JSON data into Pydantic models.
"""
import json
import os
import uuid
from pathlib import Path
from unittest.mock import patch, mock_open

import pytest
from pydantic import ValidationError
from pydantic_core import PydanticCustomError

from doc2pptx.core.models import Section, Presentation, SlideBlock
from doc2pptx.ingest.json_loader import (
    load_sections,
    load_presentation,
    _extract_sections_data,
    _adapt_section_data,
    _create_slide_block,
    _prepare_slide_content,
    _determine_content_type,
    _get_default_layout_for_type
)


# Helper function to create temp JSON files for testing
def create_temp_json_file(data, tmp_path):
    temp_file = tmp_path / "test_data.json"
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(data, f)
    return temp_file


# Test data for different scenarios
@pytest.fixture
def basic_section_data():
    return {
        "title": "Test Section",
        "type": "content",
        "content": "This is test content"
    }


@pytest.fixture
def full_section_data():
    return {
        "id": str(uuid.uuid4()),
        "title": "Full Test Section",
        "type": "bullet_list",
        "content": ["Item 1", "Item 2", "Item 3"],
        "slides": [
            {
                "id": str(uuid.uuid4()),
                "title": "Test Slide",
                "layout_name": "Titre et texte",
                "blocks": [
                    {
                        "id": str(uuid.uuid4()),
                        "title": "Block Title",
                        "content": {
                            "content_type": "bullet_points",
                            "bullet_points": ["Item 1", "Item 2", "Item 3"]
                        }
                    }
                ]
            }
        ]
    }


@pytest.fixture
def presentation_data():
    return {
        "title": "Test Presentation",
        "author": "Test Author",
        "description": "Test Description",
        "template_path": "path/to/template.pptx",
        "metadata": {"key": "value"},
        "sections": [
            {
                "title": "Section 1",
                "type": "introduction",
                "content": "Introduction content"
            },
            {
                "title": "Section 2",
                "type": "content",
                "content": ["Bullet 1", "Bullet 2"]
            }
        ]
    }


class TestLoadSections:
    def test_load_sections_from_path(self, basic_section_data, tmp_path):
        # Create a temporary JSON file
        temp_file = create_temp_json_file([basic_section_data], tmp_path)
        
        # Test loading from file path
        sections = load_sections(temp_file)
        
        assert len(sections) == 1
        assert isinstance(sections[0], Section)
        assert sections[0].title == "Test Section"
    
    def test_load_sections_from_dict(self, basic_section_data):
        # Test loading from dict with sections field
        data = {"sections": [basic_section_data]}
        sections = load_sections(data)
        
        assert len(sections) == 1
        assert sections[0].title == "Test Section"
    
    def test_load_sections_from_list(self, basic_section_data):
        # Test loading from list directly
        data = [basic_section_data]
        sections = load_sections(data)
        
        assert len(sections) == 1
        assert sections[0].title == "Test Section"
    
    def test_load_sections_with_content_field(self, basic_section_data):
        # Test loading from dict with content field instead of sections
        data = {"content": [basic_section_data]}
        sections = load_sections(data)
        
        assert len(sections) == 1
        assert sections[0].title == "Test Section"
    
    def test_load_sections_file_not_found(self):
        # Test handling of non-existent file
        with pytest.raises(FileNotFoundError):
            load_sections("nonexistent_file.json")
    
    def test_load_sections_invalid_json(self, tmp_path):
        # Create a file with invalid JSON
        path = tmp_path / "invalid.json"
        with open(path, "w") as f:
            f.write("{invalid json")
        
        with pytest.raises(json.JSONDecodeError):
            load_sections(path)
    
    def test_load_sections_unsupported_type(self):
        # Test handling of unsupported source type
        with pytest.raises(TypeError):
            load_sections(123)  # Integer is not supported
    
    def test_load_sections_invalid_structure(self):
        # Test handling of JSON without sections structure
        data = {"not_sections": "invalid"}
        
        with pytest.raises(ValueError, match="could not find sections data"):
            load_sections(data)
    
    def test_load_sections_validation_error(self):
        # Test handling of validation errors
        # Create a section with invalid type that triggers validation error
        data = [{"title": "Invalid Section", "type": "this_type_absolutely_cannot_exist"}]
        
        with pytest.raises(ValueError):
            load_sections(data)


class TestLoadPresentation:
    def test_load_presentation_from_path(self, presentation_data, tmp_path):
        # Create a temporary JSON file
        temp_file = create_temp_json_file(presentation_data, tmp_path)
        
        # Test loading from file path
        presentation = load_presentation(temp_file)
        
        assert isinstance(presentation, Presentation)
        assert presentation.title == "Test Presentation"
        assert presentation.author == "Test Author"
        assert len(presentation.sections) == 2
    
    def test_load_presentation_from_dict(self, presentation_data):
        # Test loading from dict directly
        presentation = load_presentation(presentation_data)
        
        assert isinstance(presentation, Presentation)
        assert presentation.title == "Test Presentation"
        assert presentation.author == "Test Author"
        assert len(presentation.sections) == 2
    
    def test_load_presentation_from_sections_list(self, basic_section_data):
        # Test loading from list of sections
        data = [basic_section_data, {**basic_section_data, "title": "Section 2"}]
        presentation = load_presentation(data)
        
        assert isinstance(presentation, Presentation)
        assert presentation.title == "Untitled Presentation"  # Default title
        assert len(presentation.sections) == 2
    
    def test_load_presentation_with_existing_id(self, presentation_data):
        # Test preserving existing ID
        presentation_data["id"] = "existing-id-123"
        presentation = load_presentation(presentation_data)
        
        assert presentation.id == "existing-id-123"
    
    def test_load_presentation_file_not_found(self):
        # Test handling of non-existent file
        with pytest.raises(FileNotFoundError):
            load_presentation("nonexistent_file.json")
    
    def test_load_presentation_invalid_json(self, tmp_path):
        # Create a file with invalid JSON
        path = tmp_path / "invalid.json"
        with open(path, "w") as f:
            f.write("{invalid json")
        
        with pytest.raises(json.JSONDecodeError):
            load_presentation(path)
    
    def test_load_presentation_unsupported_type(self):
        # Test handling of unsupported source type
        with pytest.raises(TypeError):
            load_presentation(123)  # Integer is not supported
    
    def test_load_presentation_invalid_section_type(self):
        # Test handling of invalid section type
        data = {
            "title": "Invalid Presentation", 
            "sections": [
                {"title": "Invalid Section", "type": "this_type_absolutely_cannot_exist"}
            ]
        }
        
        with pytest.raises(ValueError):
            load_presentation(data)


class TestExtractSectionsData:
    def test_extract_from_list(self):
        # Test extracting from a list
        data = [{"title": "Section 1"}, {"title": "Section 2"}]
        result = _extract_sections_data(data)
        
        assert result == data
    
    def test_extract_from_dict_with_sections(self):
        # Test extracting from a dict with sections field
        sections = [{"title": "Section 1"}, {"title": "Section 2"}]
        data = {"sections": sections}
        result = _extract_sections_data(data)
        
        assert result == sections
    
    def test_extract_from_dict_with_content(self):
        # Test extracting from a dict with content field
        content = [{"title": "Section 1"}, {"title": "Section 2"}]
        data = {"content": content}
        result = _extract_sections_data(data)
        
        assert result == content
    
    def test_extract_invalid_structure(self):
        # Test handling of invalid structure
        data = {"invalid": "structure"}
        
        with pytest.raises(ValueError, match="could not find sections data"):
            _extract_sections_data(data)


class TestAdaptSectionData:
    def test_adapt_minimal_section(self):
        # Test adapting a minimal section
        data = {"title": "Minimal Section"}
        result = _adapt_section_data(data)
        
        assert "id" in result  # ID should be added
        assert result["title"] == "Minimal Section"
        assert result["type"] == "content"  # Default type
        assert "slides" in result  # Slides should be added
        assert len(result["slides"]) == 1
    
    def test_adapt_section_missing_title(self):
        # Test adapting a section without title
        data = {"content": "Content without title"}
        result = _adapt_section_data(data)
        
        assert result["title"] == "Untitled Section"  # Default title
    
    def test_adapt_section_with_invalid_type(self):
        # Test handling of invalid section type
        data = {"title": "Invalid Type", "type": "this_type_absolutely_cannot_exist"}
        
        with pytest.raises(ValueError, match="Invalid section type"):
            _adapt_section_data(data)
    
    def test_adapt_section_with_type_mapping(self):
        # Test type mapping
        test_cases = [
            ("title", "title"),
            ("bullet_list", "content"),
            ("table", "content"),
            ("thank_you", "conclusion"),
        ]
        
        for input_type, expected_type in test_cases:
            data = {"title": f"Test {input_type}", "type": input_type}
            result = _adapt_section_data(data)
            assert result["type"] == expected_type
    
    def test_adapt_section_with_existing_slides(self):
        # Test adapting a section with existing slides
        data = {
            "title": "Section with Slides",
            "slides": [
                {"title": "Existing Slide"}
            ]
        }
        result = _adapt_section_data(data)
        
        assert len(result["slides"]) == 1
        assert "id" in result["slides"][0]  # ID should be added
        assert "layout_name" in result["slides"][0]  # Layout should be added
        assert "blocks" in result["slides"][0]  # Blocks should be added
    
    def test_adapt_section_without_slides(self):
        # Test adapting a section without slides
        data = {"title": "No Slides Section"}
        result = _adapt_section_data(data)
        
        assert len(result["slides"]) == 1
        assert result["slides"][0]["title"] == "No Slides Section"
    
    def test_adapt_section_with_agenda_type(self):
        # Test adapting an agenda section (special case mentioned in the code)
        data = {"title": "Agenda", "type": "agenda"}
        result = _adapt_section_data(data)
        
        assert result["type"] == "agenda"
        assert result["slides"][0]["layout_name"] == "Titre et texte"


class TestCreateSlideBlock:
    def test_create_basic_slide_block(self):
        # Test creating a basic slide block
        data = {"title": "Test Section", "content": "Test Content"}
        result = _create_slide_block(data)
        
        assert "id" in result
        assert result["title"] is None  # No subtitle in input
        assert "content" in result
        assert result["content"]["content_type"] == "text"
        assert result["content"]["text"] == "Test Content"
    
    def test_create_slide_block_with_subtitle(self):
        # Test creating a slide block with subtitle
        data = {
            "title": "Test Section", 
            "subtitle": "Test Subtitle",
            "content": "Test Content"
        }
        result = _create_slide_block(data)
        
        assert result["title"] == "Test Subtitle"


class TestPrepareSlideContent:
    def test_prepare_text_content(self):
        # Test preparing text content
        data = {"title": "Text Section", "content": "Sample text"}
        result = _prepare_slide_content(data)
        
        assert result["content_type"] == "text"
        assert result["text"] == "Sample text"
    
    def test_prepare_two_column_content(self):
        # Test preparing two-column content
        data = {
            "title": "Two Column",
            "content": {"left": "Left content", "right": "Right content"}
        }
        result = _prepare_slide_content(data)
        
        assert result["content_type"] == "text"
        assert "LEFT COLUMN" in result["text"]
        assert "RIGHT COLUMN" in result["text"]
    
    def test_prepare_bullet_points_content(self):
        # Test preparing bullet points content
        data = {
            "title": "Bullet Points",
            "type": "bullet_list",
            "content": ["Item 1", "Item 2", "Item 3"]
        }
        result = _prepare_slide_content(data)
        
        assert result["content_type"] == "bullet_points"
        assert len(result["bullet_points"]) == 3
        assert result["bullet_points"][0] == "Item 1"
    
    def test_prepare_bullet_points_from_table(self):
        # Test preparing bullet points from table rows
        data = {
            "title": "Table as Bullets",
            "type": "bullet_list",
            "content": [["Row 1, Cell 1", "Row 1, Cell 2"], ["Row 2, Cell 1", "Row 2, Cell 2"]]
        }
        result = _prepare_slide_content(data)
        
        assert result["content_type"] == "bullet_points"
        assert len(result["bullet_points"]) == 2
        assert "Row 1, Cell 1 | Row 1, Cell 2" in result["bullet_points"][0]
    
    def test_prepare_image_content(self):
        # Test preparing image content
        data = {
            "title": "Image Section",
            "type": "image_right",
            "content": "Image description",
            "image": {
                "url": "http://example.com/image.jpg",
                "alt_text": "Sample Image"
            }
        }
        result = _prepare_slide_content(data)
        
        assert result["content_type"] == "image"
        assert "image" in result
        assert result["image"]["url"] == "http://example.com/image.jpg"
        assert result["image"]["alt_text"] == "Sample Image"
        assert result["text"] == "Image description"
    
    def test_prepare_image_content_with_path(self):
        # Test preparing image content with local path
        data = {
            "title": "Image Section",
            "type": "image_left",
            "content": "Image description",
            "image": {
                "path": "/path/to/local/image.jpg",
                "alt_text": "Local Image"
            }
        }
        result = _prepare_slide_content(data)
        
        assert result["content_type"] == "image"
        assert "image" in result
        assert result["image"]["path"] == "/path/to/local/image.jpg"
    
    def test_prepare_table_content(self):
        # Test preparing table content
        data = {
            "title": "Table Section",
            "type": "table",
            "content": [
                ["Header 1", "Header 2"],
                ["Row 1, Cell 1", "Row 1, Cell 2"],
                ["Row 2, Cell 1", "Row 2, Cell 2"]
            ]
        }
        result = _prepare_slide_content(data)
        
        assert result["content_type"] == "table"
        assert "table" in result
        assert result["table"]["headers"] == ["Header 1", "Header 2"]
        assert len(result["table"]["rows"]) == 2
    
    def test_prepare_mermaid_content(self):
        # Test preparing mermaid diagram content
        data = {
            "title": "Diagram",
            "type": "chart",
            "content": "```mermaid\ngraph TD;\nA-->B;\n```"
        }
        result = _prepare_slide_content(data)
        
        assert result["content_type"] == "chart" or result["content_type"] == "mermaid"
        assert "mermaid" in result
        assert result["mermaid"]["code"] == "graph TD;\nA-->B;"
        assert result["mermaid"]["caption"] == "Diagram"


class TestDetermineContentType:
    def test_determine_from_section_type(self):
        # Test determining content type from section type
        test_cases = [
            ("bullet_list", "bullet_points"),
            ("chart", "chart"),
            ("table", "table"),
            ("image_left", "image"),
            ("image_right", "image"),
        ]
        
        for section_type, expected_content_type in test_cases:
            data = {"type": section_type}
            result = _determine_content_type(data)
            assert result == expected_content_type
    
    def test_determine_from_content_format(self):
        # Test determining content type from content format
        test_cases = [
            (["Item 1", "Item 2"], "bullet_points"),
            ("```mermaid\ngraph TD;\n```", "mermaid"),
            ("Plain text", "text"),
        ]
        
        for content, expected_content_type in test_cases:
            data = {"content": content}
            result = _determine_content_type(data)
            assert result == expected_content_type
    
    def test_determine_with_image(self):
        # Test determining content type with image present
        data = {"image": {"url": "http://example.com/image.jpg"}}
        result = _determine_content_type(data)
        
        assert result == "image"
    
    def test_determine_default(self):
        # Test determining default content type
        data = {}  # No type or content
        result = _determine_content_type(data)
        
        assert result == "text"  # Default type


class TestGetDefaultLayoutForType:
    def test_layout_mapping(self):
        # Test layout mapping for different section types
        test_cases = [
            ("title", "Diapositive de titre"),
            ("introduction", "Introduction"),
            ("content", "Titre et texte"),
            ("conclusion", "Chapitre 1"),
            ("appendix", "Titre et texte"),
            ("custom", "Titre et texte"),
            ("agenda", "Titre et texte"),
            ("unknown_type", "Titre et texte"),  # Default layout
        ]
        
        for section_type, expected_layout in test_cases:
            result = _get_default_layout_for_type(section_type)
            assert result == expected_layout, f"For section type '{section_type}', expected '{expected_layout}' but got '{result}'"