"""Shared fixtures for tests.

This module contains pytest fixtures that are shared across multiple test modules.
"""

import json
import os
import shutil
import tempfile
from pathlib import Path
from typing import Dict, List, Any, Generator, Iterator

import pytest
from _pytest.monkeypatch import MonkeyPatch
from pptx import Presentation
from pptx import Presentation as PptxPresentation

from doc2pptx.core.models import (
    ContentType,
    Section, 
    Slide, 
    SlideBlock, 
    SlideContent,
    SectionType,
    Presentation as PresentationModel
)
from doc2pptx.core.settings import Settings
from doc2pptx.ingest.json_loader import load_sections
from doc2pptx.layout.selector import LayoutSelector
from doc2pptx.ppt.template_loader import TemplateLoader
from doc2pptx.ppt.builder import PPTBuilder
from doc2pptx.ppt.overflow import OverflowHandler


@pytest.fixture
def test_env_vars(monkeypatch: MonkeyPatch) -> Iterator[Dict[str, str]]:
    """Fixture that provides test environment variables.

    Args:
        monkeypatch: Pytest monkeypatch fixture

    Yields:
        Dictionary of test environment variables
    """
    # Set test environment variables
    env_vars = {
        "OPENAI_API_KEY": "test-openai-key",
        "UNSPLASH_ACCESS_KEY": "test-unsplash-access-key",
        "UNSPLASH_SECRET_KEY": "test-unsplash-secret-key",
        "TEMPLATES_DIR": str(Path("tests") / "fixtures" / "templates"),
        "OUTPUT_DIR": str(Path("tests") / "fixtures" / "output"),
        "CACHE_DIR": str(Path("tests") / "fixtures" / "cache"),
        "DEBUG": "True",
    }

    # Apply environment variables
    for key, value in env_vars.items():
        monkeypatch.setenv(key, value)

    # Create test directories if they don't exist
    for dir_path in [env_vars["TEMPLATES_DIR"], env_vars["OUTPUT_DIR"], env_vars["CACHE_DIR"]]:
        Path(dir_path).mkdir(parents=True, exist_ok=True)

    yield env_vars

    # Clean up test directories - optionally uncomment as needed
    # for dir_path in [env_vars["OUTPUT_DIR"], env_vars["CACHE_DIR"]]:
    #     shutil.rmtree(dir_path, ignore_errors=True)


@pytest.fixture
def settings(test_env_vars: Dict[str, str]) -> Settings:
    """Fixture that provides a Settings instance with test values.

    Args:
        test_env_vars: Fixture providing test environment variables

    Returns:
        Settings instance with test values
    """
    return Settings()


@pytest.fixture
def test_resources_dir() -> Path:
    """Return the path to the test resources directory."""
    return Path(__file__).parent / "fixtures"


@pytest.fixture
def sample_json_path(test_resources_dir: Path) -> Path:
    """Return the path to the sample JSON file."""
    return test_resources_dir / "sample_input.json"


@pytest.fixture
def template_path(test_resources_dir: Path) -> Path:
    """Return the path to the template PowerPoint file."""
    return test_resources_dir / "base_template.pptx"


@pytest.fixture
def sample_json_content(sample_json_path: Path) -> Dict[str, Any]:
    """Load and return the sample JSON content."""
    with open(sample_json_path, "r", encoding="utf-8") as f:
        return json.load(f)


@pytest.fixture
def sample_sections(sample_json_content: Dict[str, Any]) -> List[Section]:
    """Load and return the sections from the sample JSON content."""
    return load_sections(sample_json_content)


@pytest.fixture
def layout_selector() -> LayoutSelector:
    """Return a layout selector instance."""
    return LayoutSelector()


@pytest.fixture
def template_loader(template_path: Path) -> TemplateLoader:
    """Return a template loader instance."""
    return TemplateLoader(str(template_path))


@pytest.fixture
def overflow_handler() -> OverflowHandler:
    """Return an overflow handler instance."""
    return OverflowHandler()


@pytest.fixture
def ppt_builder(template_path: Path, layout_selector: LayoutSelector, overflow_handler: OverflowHandler) -> PPTBuilder:
    """Return a PowerPoint builder instance."""
    return PPTBuilder(
        template_path=template_path,
        layout_selector=layout_selector,
        overflow_handler=overflow_handler
    )


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """Create a temporary directory for test files and clean it up after tests."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        yield Path(tmp_dir)


@pytest.fixture
def temp_output_pptx(temp_dir: Path) -> Path:
    """Return a path for a temporary output PowerPoint file."""
    return temp_dir / "output.pptx"


@pytest.fixture
def sample_presentation_model(sample_sections: List[Section], template_path: Path) -> PresentationModel:
    """Create a sample presentation model."""
    return PresentationModel(
        id="test-presentation",
        title="Test Presentation",
        author="Test Author",
        description="A test presentation for e2e testing",
        template_path=template_path,
        sections=sample_sections,
        metadata={
            "category": "Test",
            "keywords": "test, pytest, e2e",
        }
    )


@pytest.fixture
def simple_slide_content() -> SlideContent:
    """Create a simple text slide content for testing."""
    return SlideContent(
        content_type=ContentType.TEXT,
        text="This is a simple text content for testing purposes."
    )


@pytest.fixture
def bullet_points_slide_content() -> SlideContent:
    """Create a bullet points slide content for testing."""
    return SlideContent(
        content_type=ContentType.BULLET_POINTS,
        bullet_points=[
            "First bullet point",
            "Second bullet point",
            "Third bullet point with some additional text to make it longer",
            "Fourth bullet point"
        ]
    )


@pytest.fixture
def long_text_content() -> str:
    """Return a long text content for testing overflow handling."""
    return """
    Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam auctor, nisl nec ultricies lacinia, 
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl. Nullam auctor, nisl nec ultricies lacinia,
    nisl nisl aliquet nisl, nec ultricies nisl nisl nec nisl.
    """


@pytest.fixture
def sample_presentation() -> PresentationModel:
    """Fixture that provides a simple sample presentation for testing.

    Returns:
        Sample Presentation instance
    """
    return PresentationModel(
        id="test-simple-presentation",
        title="Sample Presentation",
        author="Test Author",
        description="This is a sample presentation for testing",
        template_path=Path("tests/fixtures/base_template.pptx"),
        sections=[
            Section(
                id="intro",
                title="Introduction",
                type=SectionType.INTRODUCTION,
                slides=[
                    Slide(
                        id="slide1",
                        title="Welcome",
                        layout_name="Title Slide",
                        blocks=[
                            SlideBlock(
                                id="title",
                                content=SlideContent(
                                    content_type=ContentType.TEXT,
                                    text="Welcome to the Presentation",
                                ),
                            ),
                            SlideBlock(
                                id="subtitle",
                                content=SlideContent(
                                    content_type=ContentType.TEXT,
                                    text="A test presentation",
                                ),
                            ),
                        ],
                    ),
                ],
            ),
            Section(
                id="content",
                title="Main Content",
                type=SectionType.CONTENT,
                slides=[
                    Slide(
                        id="slide2",
                        title="Key Points",
                        layout_name="Titre et texte",
                        blocks=[
                            SlideBlock(
                                id="content",
                                content=SlideContent(
                                    content_type=ContentType.BULLET_POINTS,
                                    bullet_points=[
                                        "Point 1",
                                        "Point 2",
                                        "Point 3",
                                    ],
                                ),
                            ),
                        ],
                    ),
                ],
            ),
        ],
    )


@pytest.fixture
def cli_runner():
    """Return a Typer CLI runner."""
    from typer.testing import CliRunner
    return CliRunner()

def pytest_configure(config):
    """Register custom markers."""
    config.addinivalue_line(
        "markers", 
        "optional: mark a test as optional, which can be skipped if certain conditions are not met"
    )

@pytest.fixture
def base_template(template_path: Path) -> Path:  # alias historique
    return template_path


@pytest.fixture
def sample_json(sample_json_path: Path) -> Path:  # alias historique
    return sample_json_path

@pytest.fixture
def tmp_output_path(temp_output_pptx: Path) -> Path:
    return temp_output_pptx

@pytest.fixture
def prs_template():
    return PptxPresentation()