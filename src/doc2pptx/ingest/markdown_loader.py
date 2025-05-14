"""
Markdown loader for doc2pptx.

This module provides functionality to load structured presentation data from
Markdown files and convert them to the Presentation model structure.
"""
import re
import logging
from pathlib import Path
from typing import Dict, List, Optional, Union, Any, Tuple
from uuid import uuid4

import mistletoe
from mistletoe.block_token import (
    Heading, Paragraph, BlockCode, Quote, List as MistletoeList,
    Table, TableRow, TableCell
)
from mistletoe.span_token import (
    RawText, Link, Image, LineBreak, Strong, Emphasis,
    Strikethrough, InlineCode
)

from doc2pptx.core.models import (
    Presentation, Section, Slide, SlideBlock, SlideContent,
    ContentType, SectionType, TableData, ImageSource
)
from doc2pptx.layout.selector import LayoutSelector

logger = logging.getLogger(__name__)

class MarkdownLoader:
    """
    Loads and converts Markdown files to Presentation models.
    
    This class handles parsing Markdown content and converting it to
    the structured Presentation model used by doc2pptx.
    """
    
    def __init__(self):
        """Initialize a Markdown loader."""
        pass
    
    def load_presentation(self, source: Union[str, Path]) -> Presentation:
        """
        Load a complete presentation from a Markdown source.
        
        This function parses the Markdown and creates a structured
        Presentation object with sections and slides.
        
        Args:
            source: Path to Markdown file or string containing Markdown
            
        Returns:
            Presentation: Validated Presentation object with sections
            
        Raises:
            FileNotFoundError: If the provided path does not exist
            ValueError: If the Markdown structure cannot be parsed
        """
        # Load Markdown content
        md_content = self._load_content(source)
        
        # Parse with mistletoe
        # CORRECTION: Ne pas utiliser 'with' avec Document
        doc = mistletoe.Document(md_content)
        
        # Extract title and metadata
        title, metadata = self._extract_metadata(doc)
        
        # Process document to create sections and slides
        sections = self._process_document(doc)
        
        # Create the presentation
        presentation = Presentation(
            id=str(uuid4()),
            title=title or "Untitled Presentation",
            sections=sections,
            **metadata
        )
        
        return presentation
    
    def _load_content(self, source: Union[str, Path]) -> str:
        """
        Load Markdown content from a file or string.
        
        Args:
            source: Path to Markdown file or string containing Markdown
            
        Returns:
            str: Markdown content
            
        Raises:
            FileNotFoundError: If the provided path does not exist
        """
        if isinstance(source, Path) or (isinstance(source, str) and "\n" not in source and len(source) < 1000):
            # Treat as file path
            path = Path(source)
            if not path.exists():
                raise FileNotFoundError(f"Markdown source file not found: {path}")
            
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                if not content:
                    raise ValueError(f"Markdown file is empty: {path}")
                return content
            except UnicodeDecodeError:
                logger.warning(f"Failed to read file with utf-8 encoding, trying with latin-1")
                with open(path, "r", encoding="latin-1") as f:
                    content = f.read()
                if not content:
                    raise ValueError(f"Markdown file is empty: {path}")
                return content
        else:
            # Treat as string content
            if not str(source).strip():
                raise ValueError("Markdown content is empty")
            return str(source)
    
    def _extract_metadata(self, document: mistletoe.Document) -> Tuple[Optional[str], Dict[str, Any]]:
        """
        Extract presentation title and metadata from document.
        
        Args:
            document: Parsed Markdown document
            
        Returns:
            Tuple[Optional[str], Dict[str, Any]]: Title and metadata dictionary
        """
        title = None
        metadata = {}
        
        # Ensure document has children
        if not document or not hasattr(document, 'children') or not document.children:
            logger.warning("Document has no content")
            return title, metadata
        
        # Look for YAML frontmatter
        frontmatter = self._extract_frontmatter(document.children)
        if frontmatter:
            try:
                import yaml
                yaml_data = yaml.safe_load(frontmatter)
                if yaml_data and isinstance(yaml_data, dict):
                    metadata.update(yaml_data)
                    if "title" in metadata:
                        title = metadata.pop("title")
            except (ImportError, yaml.YAMLError) as e:
                logger.warning(f"Error parsing YAML frontmatter: {e}")
        
        # If no title in frontmatter, look for first heading
        if not title and document.children:
            for child in document.children:
                if isinstance(child, Heading) and child.level == 1:
                    title = self._render_text_tokens(child.children)
                    break
        
        return title, metadata
    
    def _extract_frontmatter(self, children: List[Any]) -> Optional[str]:
        """
        Extract YAML frontmatter from document if present.
        
        Args:
            children: List of document child nodes
            
        Returns:
            Optional[str]: YAML frontmatter content or None
        """
        if not children:
            return None
            
        # Check if document starts with "---" for frontmatter
        frontmatter_pattern = re.compile(r"^\s*---\s*\n(.*?)\n\s*---\s*\n", re.DOTALL)
        
        # Ensure children[0] has children attribute before trying to access it
        if not hasattr(children[0], 'children'):
            return None
            
        content = self._render_text_tokens(children[0].children) if children[0].children else ""
        
        match = frontmatter_pattern.match(content)
        if match:
            return match.group(1)
            
        return None
    
    def _process_document(self, document: mistletoe.Document) -> List[Section]:
        """
        Process Markdown document and convert to sections.
        
        Args:
            document: Parsed Markdown document
            
        Returns:
            List[Section]: List of sections for the presentation
        """
        sections = []
        current_section = None
        current_slide = None
        current_blocks = []
        
        # Import rule-based layout selector
        from doc2pptx.layout.selector import LayoutSelector
        layout_selector = LayoutSelector()
        
        # Function to finalize the current slide
        def finalize_slide():
            nonlocal current_slide, current_blocks
            if current_slide and current_blocks:
                current_slide.blocks = current_blocks.copy()
                current_blocks = []
        
        # Function to finalize the current section
        def finalize_section():
            nonlocal current_section, current_slide
            if current_section:
                finalize_slide()
                if current_slide:
                    current_section.slides.append(current_slide)
                    current_slide = None
                sections.append(current_section)
        
        # Extract presentation title and metadata from first H1 if present
        presentation_title = None
        for token in document.children:
            if isinstance(token, Heading) and token.level == 1:
                presentation_title = self._render_text_tokens(token.children)
                break
        
        # Process each token in the document
        first_slide_created = False
        
        for token in document.children:
            # Skip frontmatter if present
            if isinstance(token, Paragraph) and token.children and self._is_frontmatter(token):
                continue
            
            # Process headings - determine section/slide structure
            if isinstance(token, Heading):
                level = token.level
                heading_text = self._render_text_tokens(token.children)
                
                if level == 1:
                    # H1 - Title slide (only if this is the first H1, otherwise treat as a section)
                    if not first_slide_created and heading_text == presentation_title:
                        # Create a title section
                        finalize_section()
                        current_section = Section(
                            id=str(uuid4()),
                            title=heading_text,
                            type=SectionType.TITLE,  # Set type to TITLE
                            slides=[]
                        )
                        
                        # Create a title slide with appropriate layout
                        current_slide = Slide(
                            id=str(uuid4()),
                            title=heading_text,
                            layout_name="Diapositive de titre",  # Force title slide layout
                            blocks=[]
                        )
                        first_slide_created = True
                    else:
                        # Regular H1 - New section
                        finalize_section()
                        
                        # Determine section type from content
                        section_type = SectionType.CONTENT  # Default
                        
                        # Identify section type based on title
                        if re.search(r'introduction|overview|about', heading_text, re.IGNORECASE):
                            section_type = SectionType.INTRODUCTION
                        elif re.search(r'agenda|outline|contents', heading_text, re.IGNORECASE):
                            section_type = SectionType.AGENDA
                        elif re.search(r'conclusion|summary', heading_text, re.IGNORECASE):
                            section_type = SectionType.CONCLUSION
                        elif re.search(r'appendix|reference', heading_text, re.IGNORECASE):
                            section_type = SectionType.APPENDIX
                        
                        current_section = Section(
                            id=str(uuid4()),
                            title=heading_text,
                            type=section_type,
                            slides=[]
                        )
                        
                        # Create a section header slide
                        current_slide = Slide(
                            id=str(uuid4()),
                            title=heading_text,
                            layout_name="Chapitre 1",  # Section header layout
                            blocks=[]
                        )
                
                elif level == 2:
                    # H2 - Create a new section with Chapitre 1 layout
                    # (If we're in a slide, finalize it first)
                    finalize_slide()
                    if current_slide:
                        if current_section:
                            current_section.slides.append(current_slide)
                    
                    # If no current section, create a default one
                    if not current_section:
                        current_section = Section(
                            id=str(uuid4()),
                            title=heading_text,
                            type=SectionType.SECTION_HEADER,
                            slides=[]
                        )
                    
                    # Create a section header slide
                    current_slide = Slide(
                        id=str(uuid4()),
                        title=heading_text,
                        layout_name="Chapitre 1",  # Use section header layout
                        blocks=[]
                    )
                
                elif level == 3:
                    # H3 - New slide in current section with proper layout
                    if current_section is None:
                        # Create default section if none exists
                        current_section = Section(
                            id=str(uuid4()),
                            title="Untitled Section",
                            type=SectionType.CONTENT,
                            slides=[]
                        )
                    
                    # Finalize current slide if it exists
                    finalize_slide()
                    if current_slide:
                        current_section.slides.append(current_slide)
                    
                    # Create new slide with layout based on rules
                    current_slide = Slide(
                        id=str(uuid4()),
                        title=heading_text,
                        layout_name="auto",  # Will be determined by the layout selector based on content
                        blocks=[]
                    )
                    current_blocks = []
                
                elif level >= 4:
                    # H4+ - New block in current slide with heading as title
                    if current_section is None:
                        # Create default section if none exists
                        current_section = Section(
                            id=str(uuid4()),
                            title="Untitled Section",
                            type=SectionType.CONTENT,
                            slides=[]
                        )
                    
                    if current_slide is None:
                        # Create default slide if none exists
                        current_slide = Slide(
                            id=str(uuid4()),
                            title=current_section.title,
                            layout_name="auto",  # Will be determined later
                            blocks=[]
                        )
                    
                    # Lower-level heading becomes a block title
                    # (Content will be added by subsequent tokens)
                    self.current_block_title = heading_text
            
            # Process content tokens
            elif current_section is not None:
                if current_slide is None:
                    # Create default slide if none exists
                    current_slide = Slide(
                        id=str(uuid4()),
                        title=current_section.title,
                        layout_name="auto",  # Will be determined later
                        blocks=[]
                    )
                
                # Process different content types
                block = self._create_block_from_token(token)
                if block:
                    # Add title from current H4+ heading if applicable
                    if hasattr(self, 'current_block_title') and self.current_block_title:
                        block.title = self.current_block_title
                        self.current_block_title = None
                    
                    current_blocks.append(block)
        
        # Finalize the last section
        finalize_section()
        
        # Apply layout selection based on rules
        sections = self._apply_layout_selection(sections, layout_selector)
        
        return sections

    def _apply_layout_selection(self, sections: List[Section], layout_selector: LayoutSelector) -> List[Section]:
        """
        Apply rule-based layout selection to all slides.
        
        Args:
            sections: List of sections to process
            layout_selector: Instance of LayoutSelector
            
        Returns:
            List[Section]: Updated sections with layout selection applied
        """
        for section in sections:
            for slide in section.slides:
                # Skip slides that already have explicit layouts
                if slide.layout_name and slide.layout_name != "auto":
                    continue
                    
                # Determine layout using rules
                slide.layout_name = layout_selector.get_layout_name(section, slide)
                
                # If section is Introduction, prefer the Introduction layout
                if section.type == SectionType.INTRODUCTION and slide.layout_name == "auto":
                    slide.layout_name = "Introduction"
        
        return sections
    
    def _is_frontmatter(self, token: Paragraph) -> bool:
        """
        Check if a paragraph token is YAML frontmatter.
        
        Args:
            token: Paragraph token to check
            
        Returns:
            bool: True if token is frontmatter
        """
        if not token.children:
            return False
            
        text = self._render_text_tokens(token.children)
        return text.startswith("---") and "---" in text[3:]
    
    def _create_block_from_token(self, token: Any) -> Optional[SlideBlock]:
        """
        Create a SlideBlock from a Markdown token.
        
        Args:
            token: Mistletoe token to convert
            
        Returns:
            Optional[SlideBlock]: Slide block or None if token should be skipped
        """
        block_id = str(uuid4())
        
        # Process the token based on its type
        if isinstance(token, Paragraph):
            # Regular paragraph - convert to text content
            if not hasattr(token, 'children') or not token.children:
                return None
            
            text = self._render_text_tokens(token.children)
            if not text.strip():
                return None  # Skip empty paragraphs
                
            content = SlideContent(
                content_type=ContentType.TEXT,
                text=text
            )
            
            return SlideBlock(id=block_id, content=content)
            
        elif isinstance(token, MistletoeList):
            # List - convert to bullet points
            bullet_points = self._extract_list_items(token)
            
            if not bullet_points:
                return None  # Skip empty lists
            
            # Log pour déboguer
            logger.debug(f"Processing list with token.start={token.start}, is_ordered={bool(token.start)}")
            
            content = SlideContent(
                content_type=ContentType.BULLET_POINTS,
                bullet_points=bullet_points,
                as_bullets=not token.start  # True if unordered list, False if ordered list
            )
            
            # Autre log pour confirmer
            logger.debug(f"Created SlideContent with as_bullets={not token.start}")
            
            return SlideBlock(id=block_id, content=content)
            
        elif isinstance(token, Table):
            # Table - convert to table content
            if not hasattr(token, 'header') or not token.header:
                logger.warning("Table missing header, skipping")
                return None
                
            headers, rows = self._extract_table_data(token)
            
            if not headers:
                logger.warning("Table has no headers, skipping")
                return None
                
            table_data = TableData(
                headers=headers,
                rows=rows
            )
            
            content = SlideContent(
                content_type=ContentType.TABLE,
                table=table_data
            )
            
            return SlideBlock(id=block_id, content=content)
            
        elif isinstance(token, BlockCode):
            # Code block - convert to code content
            if not hasattr(token, 'children') or not token.children:
                return None
                
            code = token.children[0].content if token.children and hasattr(token.children[0], 'content') else ""
            language = token.language or "text"
            
            content = SlideContent(
                content_type=ContentType.CODE,
                code={
                    "code": code,
                    "language": language
                }
            )
            
            return SlideBlock(id=block_id, content=content)
            
        elif isinstance(token, Quote):
            # Quote - convert to text with formatting
            if not hasattr(token, 'children') or not token.children:
                return None
                
            quote_text = self._render_text_tokens(token.children)
            
            content = SlideContent(
                content_type=ContentType.TEXT,
                text=f"> {quote_text}"
            )
            
            return SlideBlock(id=block_id, content=content)
            
        return None  # Skip other token types
    
    def _render_text_tokens(self, tokens: List[Any]) -> str:
        """
        Render a list of span tokens to text.
        
        Args:
            tokens: List of span tokens
            
        Returns:
            str: Rendered text
        """
        # Safeguard against None
        if tokens is None:
            return ""
            
        result = ""
        
        for token in tokens:
            if isinstance(token, RawText):
                result += token.content
            elif isinstance(token, LineBreak):
                result += "\n"
            elif isinstance(token, Strong):
                if hasattr(token, 'children'):
                    result += f"**{self._render_text_tokens(token.children)}**"
            elif isinstance(token, Emphasis):
                if hasattr(token, 'children'):
                    result += f"*{self._render_text_tokens(token.children)}*"
            elif isinstance(token, Strikethrough):
                if hasattr(token, 'children'):
                    result += f"~~{self._render_text_tokens(token.children)}~~"
            elif isinstance(token, InlineCode):
                if hasattr(token, 'children') and token.children:
                    result += f"`{token.children[0].content}`"
            elif isinstance(token, Link):
                if hasattr(token, 'children') and hasattr(token, 'target'):
                    text = self._render_text_tokens(token.children)
                    result += f"[{text}]({token.target})"
            elif isinstance(token, Image):
                if hasattr(token, 'children') and hasattr(token, 'src'):
                    alt = self._render_text_tokens(token.children)
                    result += f"![{alt}]({token.src})"
            elif hasattr(token, 'children'):
                result += self._render_text_tokens(token.children)
        
        return result
    
    def _extract_list_items(self, token: MistletoeList) -> List[str]:
        """
        Extract items from a list token.
        
        Args:
            token: List token
            
        Returns:
            List[str]: List of bullet point strings
        """
        items = []
        
        # Safeguard against None or empty children
        if not hasattr(token, 'children') or not token.children:
            return items
        
        for item in token.children:
            # Safeguard against None children
            if not hasattr(item, 'children') or not item.children:
                continue
                
            # Get the text content of the list item
            text = ""
            for child in item.children:
                if isinstance(child, Paragraph):
                    if hasattr(child, 'children'):
                        text += self._render_text_tokens(child.children)
                elif isinstance(child, MistletoeList):
                    # Handle nested lists - add indentation
                    if hasattr(child, 'children'):
                        nested_items = self._extract_list_items(child)
                        if text:
                            items.append(text)
                        for nested_item in nested_items:
                            items.append(f"    {nested_item}")
                        text = ""
                else:
                    if hasattr(child, 'children'):
                        text += self._render_text_tokens(child.children)
            
            if text:
                items.append(text)
        
        return items
        
    def _extract_table_data(self, token: Table) -> Tuple[List[str], List[List[str]]]:
        """
        Extract headers and rows from a table token.
        
        Args:
            token: Table token
            
        Returns:
            Tuple[List[str], List[List[str]]]: Headers and rows
        """
        headers = []
        rows = []
        
        # Extract headers from the first row
        if hasattr(token, 'header') and token.header and hasattr(token.header, 'children'):
            for cell in token.header.children:
                if hasattr(cell, 'children'):
                    cell_text = self._render_text_tokens(cell.children)
                    headers.append(cell_text)
        
        # Extract data rows
        if hasattr(token, 'children'):
            for row in token.children:
                if hasattr(row, 'children'):
                    row_data = []
                    for cell in row.children:
                        if hasattr(cell, 'children'):
                            cell_text = self._render_text_tokens(cell.children)
                            row_data.append(cell_text)
                    
                    # Ensure row has the same length as headers
                    if len(row_data) < len(headers):
                        # Pad with empty strings
                        row_data.extend([''] * (len(headers) - len(row_data)))
                    elif len(row_data) > len(headers):
                        # Truncate if too many cells
                        row_data = row_data[:len(headers)]
                    
                    rows.append(row_data)
        
        # Clean headers and rows to handle any empty values
        headers = [h if h else '' for h in headers]
        for i, row in enumerate(rows):
            rows[i] = [str(cell) if cell is not None else '' for cell in row]
        
        # Log extracted data for debugging
        logger.debug(f"Extracted table with {len(headers)} headers and {len(rows)} rows")
        if headers:
            logger.debug(f"Headers: {headers}")
        if rows:
            logger.debug(f"First row sample: {rows[0] if rows else []}")
        
        return headers, rows
    
    def detect_image_references(self, md_content: str) -> List[Dict[str, str]]:
        """
        Detect image references in Markdown content.
        
        Args:
            md_content: Markdown content string
            
        Returns:
            List[Dict[str, str]]: List of image references with src and alt
        """
        # Safeguard against None
        if md_content is None:
            return []
            
        # Pattern for Markdown image syntax: ![alt text](image_url)
        image_pattern = re.compile(r'!\[(.*?)\]\((.*?)\)')
        
        images = []
        for match in image_pattern.finditer(md_content):
            alt_text, src = match.groups()
            images.append({
                "alt_text": alt_text,
                "src": src
            })
        
        return images


def load_presentation_from_markdown(source: Union[str, Path]) -> Presentation:
    """
    Load a presentation from a Markdown source.
    
    This function is a convenience wrapper around MarkdownLoader.
    
    Args:
        source: Path to Markdown file or string containing Markdown
        
    Returns:
        Presentation: The loaded and parsed presentation
    
    Raises:
        FileNotFoundError: If the provided path does not exist
        ValueError: If the Markdown structure cannot be parsed
    """
    loader = MarkdownLoader()
    return loader.load_presentation(source)