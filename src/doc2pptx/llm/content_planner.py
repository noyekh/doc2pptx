# src/doc2pptx/llm/content_planner.py

import json
import logging
from typing import Dict, List, Optional, Any, Union, Tuple

from doc2pptx.core.models import Section, Slide, SlideBlock, ContentType
from doc2pptx.llm.optimizer import PresentationOptimizer
from doc2pptx.ppt.template_loader import TemplateInfo

logger = logging.getLogger(__name__)

class ContentPlanner:
    """
    Plans the distribution of content across slides for optimal presentation.
    
    This class analyzes section content and determines the optimal number of slides,
    layouts, and content distribution for effective presentation.
    """
    
    def __init__(self, optimizer: Optional[PresentationOptimizer] = None):
        """
        Initialize a content planner.
        
        Args:
            optimizer: Optional PresentationOptimizer. If not provided, a new one will be created.
        """
        self.optimizer = optimizer or PresentationOptimizer()
        self.use_ai = bool(self.optimizer.client)
        
        if not self.use_ai:
            logger.warning("AI client not available. Content planning will use simple heuristics.")
    
    def plan_section_content(self, section: Section, 
                            template_info: Optional[TemplateInfo] = None,
                            max_content_per_slide: int = 2000) -> Section:
        """
        Plan the optimal distribution of content for a section.
        
        This method analyzes the section content and determines the optimal number of slides,
        layouts, and content distribution. It modifies the section in place and returns it.
        
        Args:
            section: The section to plan.
            template_info: Optional template information for layout selection.
            max_content_per_slide: Maximum content length per slide (used as fallback).
            
        Returns:
            The modified section with optimized slides.
        """
        if not self.use_ai:
            # Fallback to simple heuristic-based planning
            return self._plan_section_heuristic(section, template_info, max_content_per_slide)
        
        # Try AI-based planning
        try:
            return self._plan_section_with_ai(section, template_info)
        except Exception as e:
            logger.error(f"Error in AI-based content planning: {e}. Falling back to heuristics.")
            return self._plan_section_heuristic(section, template_info, max_content_per_slide)

    def _plan_section_with_ai(self, section: Section, 
                            template_info: Optional[TemplateInfo] = None) -> Section:
        """
        Plan section content distribution using AI.
        
        Args:
            section: The section to plan.
            template_info: Optional template information for layout selection.
            
        Returns:
            The modified section with optimized slides.
        """
        # Extract available layouts information
        layouts_info = self._extract_layouts_info(template_info)
        
        # Prepare content for analysis
        section_content = self._extract_section_content(section)
        
        # Create the prompt for the AI
        prompt = self._create_planning_prompt(section, section_content, layouts_info)
        
        # Validate that we have a client
        if not self.optimizer or not self.optimizer.client:
            logger.warning("AI client not available for content planning")
            raise ValueError("AI client not available")
        
        # Get AI response
        try:
            response = self.optimizer.client.chat.completions.create(
                model=self.optimizer.model,
                messages=[
                    {"role": "system", "content": "You are a PowerPoint design expert assistant."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                response_format={"type": "json_object"}
            )
            
            # Parse and validate the response
            response_text = response.choices[0].message.content.strip()
            content_plan = json.loads(response_text)
            
            # Validate the structure
            if not isinstance(content_plan, dict) or "slides" not in content_plan:
                logger.warning("Invalid content plan structure returned by AI")
                logger.debug(f"Raw content plan: {response_text}")
                raise ValueError("Invalid content plan structure")
            
            # Apply the plan to the section
            new_slides = self._apply_content_plan(section, content_plan, template_info)
            
            # Update the section with new slides
            updated_section = section.model_copy(deep=True)
            updated_section.slides = new_slides
            
            logger.info(f"Successfully planned content for section '{section.title}' with AI")
            return updated_section
            
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse AI response: {e}")
            raise ValueError(f"Failed to parse AI response: {e}")
        except Exception as e:
            logger.error(f"Error in AI content planning: {e}")
            raise
    
    def _extract_layouts_info(self, template_info: Optional[TemplateInfo]) -> Dict[str, Any]:
        """
        Extract relevant layout information for AI planning.
        
        Args:
            template_info: Template information containing layout details.
            
        Returns:
            Dictionary with layout information.
        """
        layouts_info = {}
        
        if template_info:
            for layout_name, layout_info in template_info.layout_map.items():
                layouts_info[layout_name] = {
                    "supports_title": layout_info.supports_title,
                    "supports_content": layout_info.supports_content,
                    "supports_table": layout_info.supports_table,
                    "supports_image": layout_info.supports_image,
                    "supports_chart": layout_info.supports_chart,
                    "max_content_blocks": layout_info.max_content_blocks,
                    "description": getattr(layout_info, "ai_description", "")
                }
        else:
            # Provide some default layouts for AI to consider
            layouts_info = {
                "Title Slide": {
                    "supports_title": True,
                    "supports_content": False,
                    "supports_table": False,
                    "supports_image": False,
                    "supports_chart": False,
                    "max_content_blocks": 0,
                    "description": "Main title slide with optional subtitle"
                },
                "Title and Content": {
                    "supports_title": True,
                    "supports_content": True,
                    "supports_table": False,
                    "supports_image": False,
                    "supports_chart": False,
                    "max_content_blocks": 1,
                    "description": "Slide with title and main content area"
                },
                "Title and Two Content": {
                    "supports_title": True,
                    "supports_content": True,
                    "supports_table": False,
                    "supports_image": False,
                    "supports_chart": False,
                    "max_content_blocks": 2,
                    "description": "Slide with title and two content columns"
                },
                "Title and Table": {
                    "supports_title": True,
                    "supports_content": True,
                    "supports_table": True,
                    "supports_image": False,
                    "supports_chart": False,
                    "max_content_blocks": 2,
                    "description": "Slide with title, text content and table"
                }
            }
        
        return layouts_info
    
    def _extract_section_content(self, section: Section) -> Dict[str, Any]: 
        """
        Extract content from a section for AI analysis.
        
        Args:
            section: The section to extract content from.
            
        Returns:
            Dictionary with extracted content.
        """
        content = {
            "title": section.title,
            "type": section.type.value,
            "description": section.description,
            "content_blocks": []
        }
        
        # Extract content from existing slides
        for slide in section.slides:
            slide_content = {
                "title": slide.title,
                "blocks": []
            }
            
            for block in slide.blocks:
                block_content = {
                    "title": block.title,
                    "content_type": block.content.content_type.value,
                }
                
                # Extract specific content based on type
                if block.content.content_type == ContentType.TEXT and block.content.text:
                    block_content["text"] = block.content.text
                elif block.content.content_type == ContentType.BULLET_POINTS and block.content.bullet_points:
                    block_content["bullet_points"] = block.content.bullet_points
                elif block.content.content_type == ContentType.TABLE and block.content.table:
                    table = block.content.table
                    block_content["table"] = {
                        "headers": table.headers,
                        "row_count": len(table.rows),
                        # Conserver aussi les données des lignes
                        "rows": table.rows
                    }
                elif block.content.content_type == ContentType.IMAGE and block.content.image:
                    block_content["image"] = {
                        "description": "Image content"
                    }
                
                slide_content["blocks"].append(block_content)
            
            content["content_blocks"].append(slide_content)
        
        return content
    
    def _create_planning_prompt(self, section: Section, 
                            section_content: Dict[str, Any],
                            layouts_info: Dict[str, Any]) -> str:
        """
        Create a prompt for AI content planning.
        
        Args:
            section: The section to plan.
            section_content: Extracted section content.
            layouts_info: Available layout information.
            
        Returns:
            Prompt string for AI.
        """
        return f"""
        As a PowerPoint design expert, plan the optimal distribution of this content across multiple slides.
        
        SECTION INFORMATION:
        {json.dumps(section_content, indent=2)}
        
        AVAILABLE LAYOUTS:
        {json.dumps(layouts_info, indent=2)}
        
        Consider these guidelines:
        1. Each slide should have a logical amount of content (not too much, not too little).
        2. Content should be distributed based on themes and logical breaks.
        3. Choose appropriate layouts based on content type and structure.
        4. Keep related content together when possible.
        5. For text-heavy content, consider using bullet points for better readability.
        6. IMPORTANT: Tables should preferably be accompanied by related text content. Don't put tables alone on slides if there is relevant text nearby.
        7. Aim for visual balance and readability.
        8. IMPORTANT: Always keep introductory text and their associated lists together on the same slide.
        For example, keep "Strategy" text and its bullet points together, or "Content Pillars" and its numbered list.
        Avoid separating a heading from its immediately following content.
        9. IMPORTANT: Preserve the format of numbered lists (1, 2, 3, etc.) versus bullet point lists.
        If the original content uses numbers, maintain the numbered format in your plan.
        10. Balance slide content - avoid slides that appear too empty or too crowded.
        
        Return a detailed content plan in this JSON format:
        {{
            "slides": [
                {{
                    "title": "Slide Title",
                    "layout": "Layout Name",
                    "content": [
                        {{
                            "type": "text|bullet_points|table|image",
                            "content": "Text content or bullet points array or table object",
                            "notes": "Optional explanation of why this content is placed here",
                            "is_numbered": true/false  // Include this for bullet_points type to indicate if it should be numbered
                        }}
                    ]
                }}
            ],
            "recommendations": "Optional overall recommendations for the section"
        }}
        
        The "content" field should contain the actual content text for text and bullet points,
        and for tables, use the format: {{"headers": [...], "rows": [...]}}
        """
    
    def _apply_content_plan(self, section: Section, 
                        content_plan: Dict[str, Any],
                        template_info: Optional[TemplateInfo]) -> List[Slide]:
        """
        Apply a content plan to create optimized slides.
        
        Args:
            section: The original section.
            content_plan: The AI-generated content plan.
            template_info: Optional template information.
            
        Returns:
            List of new slides based on the content plan.
        """
        new_slides = []
        
        # Process each slide in the plan
        for slide_plan in content_plan.get("slides", []):
            slide_title = slide_plan.get("title", section.title)
            layout_name = slide_plan.get("layout", "Title and Content")
            
            # Map generic layout names to template-specific ones if needed
            if template_info:
                layout_name = self._map_generic_layout_to_template(layout_name, template_info)
            
            # Create blocks from the content
            blocks = []
            for content_item in slide_plan.get("content", []):
                content_type = content_item.get("type", "text")
                content_data = content_item.get("content", "")
                is_numbered = content_item.get("is_numbered", False)
                
                # Create the appropriate content object
                block = self._create_block_from_content(content_type, content_data, is_numbered)
                if block:
                    # If content has a title specified, use it
                    block_title = content_item.get("title")
                    if block_title:
                        block.title = block_title
                    blocks.append(block)
            
            # Create the slide
            from uuid import uuid4
            slide = Slide(
                id=str(uuid4()),
                title=slide_title,
                layout_name=layout_name,
                blocks=blocks,
                notes=slide_plan.get("notes", None)
            )
            
            new_slides.append(slide)
        
        return new_slides
    
    def _map_generic_layout_to_template(self, generic_layout: str, 
                                       template_info: TemplateInfo) -> str:
        """
        Map a generic layout name to a template-specific one.
        
        Args:
            generic_layout: Generic layout name (e.g., "Title and Content").
            template_info: Template information with available layouts.
            
        Returns:
            Template-specific layout name.
        """
        # Common mapping patterns
        mappings = {
            "Title Slide": ["titre", "title", "diapositive de titre"],
            "Title and Content": ["titre et texte", "content", "text"],
            "Title and Two Content": ["titre et 3 colonnes", "columns", "two content"],
            "Title and Table": ["titre et tableau", "table"],
            "Title Only": ["titre seul", "chapitre", "section"],
            "Title and Image": ["titre et texte 1 visuel", "image"],
            "Title and Chart": ["titre et texte 1 histogramme", "chart", "graph"]
        }
        
        # Check for exact matches first
        if generic_layout in template_info.layout_map:
            return generic_layout
        
        # Try case-insensitive matching
        layout_lower = generic_layout.lower()
        for template_layout in template_info.layout_map.keys():
            if template_layout.lower() == layout_lower:
                return template_layout
        
        # Try mapping based on patterns
        for template_layout in template_info.layout_map.keys():
            template_lower = template_layout.lower()
            
            # Check if the template layout matches any of the patterns for this generic layout
            for generic, patterns in mappings.items():
                if generic.lower() == layout_lower:
                    for pattern in patterns:
                        if pattern in template_lower:
                            return template_layout
        
        # Default to the first compatible layout based on content needs
        if "table" in layout_lower and template_info.table_layouts:
            return template_info.table_layouts[0]
        elif "chart" in layout_lower and template_info.chart_layouts:
            return template_info.chart_layouts[0]
        elif "image" in layout_lower and template_info.image_layouts:
            return template_info.image_layouts[0]
        elif "two" in layout_lower or "column" in layout_lower and template_info.two_content_layouts:
            return template_info.two_content_layouts[0]
        elif "content" in layout_lower and template_info.content_layouts:
            return template_info.content_layouts[0]
        elif "title" in layout_lower and "only" in layout_lower and template_info.title_layouts:
            return template_info.title_layouts[0]
        
        # Fallback to the first layout if no match found
        logger.warning(f"Could not map generic layout '{generic_layout}' to a template layout. Using first available.")
        return next(iter(template_info.layout_map.keys()))
    
    def _create_block_from_content(self, content_type: str, content_data: Any, is_numbered: bool = False) -> Optional[SlideBlock]:
        """
        Create a SlideBlock from content specification.
        
        Args:
            content_type: Type of content (text, bullet_points, table, image).
            content_data: The content data.
            is_numbered: Whether bullet_points should be rendered as a numbered list.
            
        Returns:
            SlideBlock instance or None if creation failed.
        """
        from uuid import uuid4
        from doc2pptx.core.models import SlideBlock, SlideContent, ContentType, TableData
        
        try:
            # Create content based on type
            content = None
            
            if content_type == "text":
                # For text content
                if isinstance(content_data, str):
                    content = SlideContent(
                        content_type=ContentType.TEXT,
                        text=content_data
                    )
            
            elif content_type == "bullet_points":
                # For bullet point content
                if isinstance(content_data, list):
                    content = SlideContent(
                        content_type=ContentType.BULLET_POINTS,
                        bullet_points=content_data,
                        as_bullets=not is_numbered  # False for numbered list, True for bullet points
                    )
                elif isinstance(content_data, str):
                    # Split string into bullet points
                    bullet_points = [line.strip() for line in content_data.split('\n') if line.strip()]
                    content = SlideContent(
                        content_type=ContentType.BULLET_POINTS,
                        bullet_points=bullet_points,
                        as_bullets=not is_numbered  # False for numbered list, True for bullet points
                    )
            
            elif content_type == "table":
                # For table content
                if isinstance(content_data, dict) and "headers" in content_data:
                    # Utiliser les données de lignes si présentes, sinon générer à partir de row_count
                    if "rows" in content_data and isinstance(content_data["rows"], list) and content_data["rows"]:
                        # Utiliser directement les données réelles des lignes
                        table_data = TableData(
                            headers=content_data["headers"],
                            rows=content_data["rows"]
                        )
                    elif "row_count" in content_data and isinstance(content_data["row_count"], int):
                        # Générer des données génériques basées sur les en-têtes et row_count
                        # (fallback pour la compatibilité)
                        generic_rows = []
                        for i in range(content_data["row_count"]):
                            row = []
                            for header in content_data["headers"]:
                                # Enlever le dernier header s'il contient 'style:'
                                if isinstance(header, str) and header.startswith("style:"):
                                    continue
                                # Créer une valeur générique basée sur le header
                                first_word = header.split()[0] if isinstance(header, str) else "Item"
                                row.append(f"{first_word} {i+1}")
                            generic_rows.append(row)
                        
                        table_data = TableData(
                            headers=content_data["headers"],
                            rows=generic_rows
                        )
                    else:
                        logger.warning("Table content missing required row data")
                        return None
                    
                    content = SlideContent(
                        content_type=ContentType.TABLE,
                        table=table_data
                    )
            
            elif content_type == "image":
                # For image content - simplified implementation
                from doc2pptx.core.models import ImageSource
                content = SlideContent(
                    content_type=ContentType.IMAGE,
                    image=ImageSource(
                        alt_text="Image",
                        query="placeholder image"  # Placeholder query
                    )
                )
            
            # Create the block if we have valid content
            if content:
                return SlideBlock(
                    id=str(uuid4()),
                    content=content
                )
            
            return None
            
        except Exception as e:
            logger.error(f"Error creating block from content: {e}")
            return None
            
    def _find_related_content_for_table(self, table_block: SlideBlock, all_content: List[SlideBlock], 
                                    used_blocks: List[SlideBlock]) -> Tuple[List[SlideBlock], Optional[str]]:
        """
        Find related text content for a table block.
        """
        logger.debug(f"=== Finding related content for table block ===")
        related_blocks = []
        title_candidate = table_block.title
        logger.debug(f"Initial title_candidate from block.title: '{title_candidate}'")
        
        # Get the position of the table block in the content
        try:
            table_pos = all_content.index(table_block)
            logger.debug(f"Table position in content: {table_pos+1} of {len(all_content)}")
        except ValueError:
            logger.debug(f"Table block not found in content list. Generating title directly.")
            # Fallback: Generate title from table content if no position found
            if not title_candidate and table_block.content and table_block.content.table:
                title_candidate = self._generate_title_from_table_data(table_block.content.table)
                logger.debug(f"Generated title from table data: '{title_candidate}'")
            return related_blocks, title_candidate
        
        # First, try to find a text block with the same title
        if table_block.title:
            logger.debug(f"Looking for text block with same title: '{table_block.title}'")
            for i, block in enumerate(all_content):
                if (block not in used_blocks and block != table_block and
                    block.content and block.content.content_type in [ContentType.TEXT, ContentType.BULLET_POINTS] and
                    block.title and block.title == table_block.title):
                    logger.debug(f"Found matching text block at position {i+1}")
                    related_blocks.append(block)
                    used_blocks.append(block)
                    break
        
        # If no match by title, look for text blocks immediately around the table (2 positions before and 1 after)
        if not related_blocks:
            logger.debug("No text block with matching title found. Looking for adjacent blocks...")
            
            # Look before the table (up to 2 positions)
            for offset in range(1, 3):
                if table_pos - offset >= 0:
                    prev_block = all_content[table_pos - offset]
                    logger.debug(f"Checking block {table_pos - offset + 1}: type={prev_block.content.content_type if prev_block.content else 'None'}, title='{prev_block.title if hasattr(prev_block, 'title') else 'None'}'")
                    if (prev_block not in used_blocks and
                        prev_block.content and prev_block.content.content_type in [ContentType.TEXT, ContentType.BULLET_POINTS]):
                        # Get title from block if it has one and we don't
                        if prev_block.title and not title_candidate:
                            title_candidate = prev_block.title
                            logger.debug(f"Using title from previous block: '{title_candidate}'")
                        # Try to extract title from text content if no title available
                        elif not title_candidate and prev_block.content.content_type == ContentType.TEXT and prev_block.content.text:
                            # Extract first line or sentence as title
                            text = prev_block.content.text
                            first_line = text.split('\n', 1)[0].strip()
                            logger.debug(f"First line from text block: '{first_line[:50]}{'...' if len(first_line) > 50 else ''}'")
                            if len(first_line) > 5 and len(first_line) < 100:  # Reasonable title length
                                title_candidate = first_line
                                logger.debug(f"Using first line as title: '{title_candidate}'")
                        
                        related_blocks.append(prev_block)
                        used_blocks.append(prev_block)
                        logger.debug(f"Added previous block at position {table_pos - offset + 1} to related blocks")
                        break  # Only get one block before
            
            # Look after the table (just 1 position)
            if table_pos + 1 < len(all_content):
                next_block = all_content[table_pos + 1]
                logger.debug(f"Checking block after table: type={next_block.content.content_type if next_block.content else 'None'}, title='{next_block.title if hasattr(next_block, 'title') else 'None'}'")
                if (next_block not in used_blocks and
                    next_block.content and next_block.content.content_type in [ContentType.TEXT, ContentType.BULLET_POINTS]):
                    if next_block.title and not title_candidate:
                        title_candidate = next_block.title
                        logger.debug(f"Using title from next block: '{title_candidate}'")
                    related_blocks.append(next_block)
                    used_blocks.append(next_block)
                    logger.debug(f"Added next block at position {table_pos + 2} to related blocks")
        
        # If still no title candidate, generate from table content
        if not title_candidate and table_block.content and table_block.content.table:
            logger.debug("No title candidate found from related blocks. Generating from table data.")
            title_candidate = self._generate_title_from_table_data(table_block.content.table)
            logger.debug(f"Generated title from table data: '{title_candidate}'")
        
        logger.debug(f"Returning title_candidate: '{title_candidate}' and {len(related_blocks)} related blocks")
        return related_blocks, title_candidate
    
    def _generate_title_from_table_data(self, table_data) -> str:
        """
        Generate a meaningful title from table data.
        """
        logger.debug("=== Generating title from table data ===")
        
        if not hasattr(table_data, 'headers') or not table_data.headers:
            logger.debug("Table has no headers. Using default title.")
            return "Tableau de données"
        
        logger.debug(f"Original headers: {table_data.headers}")
        
        # Clean headers (remove any "style:" header)
        headers = [h for h in table_data.headers if not (isinstance(h, str) and h.startswith("style:"))]
        
        logger.debug(f"Cleaned headers: {headers}")
        
        if not headers:
            logger.debug("No valid headers after cleaning. Using default title.")
            return "Tableau de données"
        
        # Use the first header as main subject
        subject = headers[0]
        
        # If there are 2-3 headers, create a more descriptive title
        if len(headers) == 2:
            title = f"Données de {subject} et {headers[1]}"
            logger.debug(f"Created title with 2 headers: '{title}'")
            return title
        elif len(headers) == 3:
            title = f"Tableau de {subject}, {headers[1]} et {headers[2]}"
            logger.debug(f"Created title with 3 headers: '{title}'")
            return title
        elif len(headers) > 3:
            title = f"Tableau de {subject} et autres données"
            logger.debug(f"Created title with >3 headers: '{title}'")
            return title
        else:
            title = f"Données de {subject}"
            logger.debug(f"Created title with 1 header: '{title}'")
            return title
    
    def _plan_section_heuristic(self, section: Section, 
                        template_info: Optional[TemplateInfo] = None,
                        max_content_per_slide: int = 3000) -> Section:
        """
        Plan section content distribution using simple heuristics.
        This is a fallback when AI planning is not available.
        """
        # Log initial state
        logger.debug(f"=== Starting content planning for section: '{section.title}' ===")
        logger.debug(f"Section has {len(section.slides)} slides and {sum(len(slide.blocks) for slide in section.slides)} content blocks")
        
        # Si section already has slides with content, preserve them
        if section.slides and all(slide.blocks for slide in section.slides):
            logger.debug("Section already has slides with content, preserving them")
            # Log slide titles and content types
            for i, slide in enumerate(section.slides):
                has_table = any(block.content and block.content.content_type == ContentType.TABLE for block in slide.blocks)
                logger.debug(f"Slide {i+1}: title='{slide.title}', layout='{slide.layout_name}', has_table={has_table}, blocks={len(slide.blocks)}")
            return section
        
        # Collect all content from the section
        all_content = []
        for slide in section.slides:
            for block in slide.blocks:
                all_content.append(block)
        
        # If no content, return the original section
        if not all_content:
            logger.debug("Section has no content, returning original section")
            return section
        
        # Group content by type
        text_blocks = []
        table_blocks = []
        image_blocks = []
        other_blocks = []
        
        for block in all_content:
            if not block.content:
                continue
                
            if block.content.content_type == ContentType.TEXT or block.content.content_type == ContentType.BULLET_POINTS:
                text_blocks.append(block)
            elif block.content.content_type == ContentType.TABLE:
                table_blocks.append(block)
            elif block.content.content_type == ContentType.IMAGE:
                image_blocks.append(block)
            else:
                other_blocks.append(block)
        
        logger.debug(f"Content blocks: text={len(text_blocks)}, table={len(table_blocks)}, image={len(image_blocks)}, other={len(other_blocks)}")
        
        # Create optimized slides
        new_slides = []
        used_blocks = []  # Track blocks that have been used
        
        # DEBUG: Log table details
        for i, table_block in enumerate(table_blocks):
            if table_block.content and table_block.content.table:
                table_data = table_block.content.table
                logger.debug(f"Table {i+1} details: block_title='{table_block.title}', headers={table_data.headers if hasattr(table_data, 'headers') else None}, rows={len(table_data.rows) if hasattr(table_data, 'rows') and table_data.rows else 0}")
        
        # Process tables - each table goes with related text content
        logger.debug(f"--- Processing {len(table_blocks)} table blocks ---")
        for i, table_block in enumerate(table_blocks):
            if table_block in used_blocks:
                logger.debug(f"Table block {i+1} already used, skipping")
                continue
                
            from uuid import uuid4
            
            logger.debug(f"Processing table block {i+1}")
            
            # Find related text content for this table
            related_blocks, suggested_title = self._find_related_content_for_table(
                table_block, all_content, used_blocks
            )
            
            logger.debug(f"_find_related_content_for_table returned: suggested_title='{suggested_title}', related_blocks={len(related_blocks)}")
            
            # Create blocks for the slide, with table first then text content
            slide_blocks = [table_block]
            used_blocks.append(table_block)
            
            # Add related blocks
            slide_blocks.extend(related_blocks)
            
            # Create the slide
            actual_title = suggested_title or table_block.title or section.title
            logger.debug(f"Creating slide for table with title: '{actual_title}' (from suggested='{suggested_title}', block_title='{table_block.title}', section_title='{section.title}')")
            
            # Get the layout
            layout_name = self._select_layout_for_content("table", template_info)
            logger.debug(f"Selected layout for table: '{layout_name}'")
            
            slide = Slide(
                id=str(uuid4()),
                title=actual_title,
                layout_name=layout_name,
                blocks=slide_blocks
            )
            
            logger.debug(f"Created new slide: id={slide.id}, title='{slide.title}', layout='{slide.layout_name}', blocks={len(slide.blocks)}")
            new_slides.append(slide)
        
        # Get remaining text blocks that haven't been used with tables
        remaining_text_blocks = [block for block in text_blocks if block not in used_blocks]
        
        # Identify content groups that should stay together
        content_groups = {}
        group_counter = 0
        
        for i, block in enumerate(remaining_text_blocks):
            # Check if this block starts a logical group (based on title or content)
            starts_group = False
            
            if block.title:
                lower_title = block.title.lower()
                if any(keyword in lower_title for keyword in ["stratégie", "piliers", "introduction", "résumé"]):
                    starts_group = True
            
            # Check text content for group indicators
            if block.content.content_type == ContentType.TEXT and block.content.text:
                text_lower = block.content.text.lower()
                if any(keyword in text_lower for keyword in ["notre stratégie", "piliers de", "principaux objectifs"]):
                    starts_group = True
            
            if starts_group:
                group_id = f"group_{group_counter}"
                group_counter += 1
                content_groups[i] = group_id
                
                # Mark subsequent bullet points as part of this group
                for j in range(i+1, len(remaining_text_blocks)):
                    next_block = remaining_text_blocks[j]
                    # Stop if we hit another potential group starter
                    if next_block.title and len(next_block.title) > 5:
                        break
                        
                    # Include bullet points in this group
                    if next_block.content.content_type == ContentType.BULLET_POINTS:
                        content_groups[j] = group_id
                    # Include short text paragraphs too
                    elif next_block.content.content_type == ContentType.TEXT:
                        if next_block.content.text and len(next_block.content.text) < 100:
                            content_groups[j] = group_id
                        else:
                            break
        
        # Process text blocks - group them by size and content relationships
        current_blocks = []
        current_size = 0
        current_group = None
        current_slide_title = None
        
        for i, text_block in enumerate(remaining_text_blocks):
            # Check if this block is part of a content group
            block_group = content_groups.get(i)
            
            # Estimate block size
            block_size = 0
            if text_block.content.content_type == ContentType.TEXT and text_block.content.text:
                block_size = len(text_block.content.text)
            elif text_block.content.content_type == ContentType.BULLET_POINTS and text_block.content.bullet_points:
                block_size = sum(len(point) for point in text_block.content.bullet_points)
            
            # If this block has a title, it might become the slide title
            if text_block.title and (not current_slide_title or len(current_blocks) == 0):
                current_slide_title = text_block.title
            
            # Start new slide logic:
            # 1. If adding this block would exceed the limit AND
            # 2. It's not part of the current group (or there is no current group) AND
            # 3. We already have blocks to put on a slide
            if (current_blocks and 
                current_size + block_size > max_content_per_slide and 
                (block_group is None or block_group != current_group)):
                
                from uuid import uuid4
                slide = Slide(
                    id=str(uuid4()),
                    title=current_slide_title or section.title,
                    layout_name=self._select_layout_for_content("text", template_info),
                    blocks=current_blocks.copy()
                )
                new_slides.append(slide)
                current_blocks = []
                current_size = 0
                current_slide_title = text_block.title  # Reset title for new slide
            
            # Update current group
            if block_group is not None:
                current_group = block_group
            
            # Add block to current collection
            current_blocks.append(text_block)
            current_size += block_size
        
        # Add any remaining text blocks
        if current_blocks:
            from uuid import uuid4
            slide = Slide(
                id=str(uuid4()),
                title=current_slide_title or section.title,
                layout_name=self._select_layout_for_content("text", template_info),
                blocks=current_blocks
            )
            new_slides.append(slide)
        
        # Process image blocks - try to pair with unused text blocks
        remaining_text_blocks = [block for block in text_blocks if block not in used_blocks]
        
        for image_block in image_blocks:
            if image_block in used_blocks:
                continue
                
            # Try to find related text for this image
            found_text = False
            related_text_block = None
            
            # Look for text blocks with the same title
            if image_block.title:
                for text_block in remaining_text_blocks:
                    if text_block.title and text_block.title == image_block.title:
                        related_text_block = text_block
                        remaining_text_blocks.remove(text_block)
                        found_text = True
                        break
            
            # Try to find a text block near the image in the content list
            if not found_text:
                try:
                    img_pos = all_content.index(image_block)
                    # Check blocks before and after image
                    for pos in [img_pos - 1, img_pos + 1]:
                        if 0 <= pos < len(all_content):
                            block = all_content[pos]
                            if (block in remaining_text_blocks and
                                block.content and block.content.content_type in [ContentType.TEXT, ContentType.BULLET_POINTS]):
                                related_text_block = block
                                remaining_text_blocks.remove(block)
                                found_text = True
                                break
                except ValueError:
                    pass
            
            # Create a slide with the image and related text
            slide_blocks = [image_block]
            used_blocks.append(image_block)
            
            if related_text_block:
                slide_blocks.append(related_text_block)
                used_blocks.append(related_text_block)
            
            # Determine slide title
            slide_title = None
            if image_block.title:
                slide_title = image_block.title
            elif related_text_block and related_text_block.title:
                slide_title = related_text_block.title
            
            from uuid import uuid4
            slide = Slide(
                id=str(uuid4()),
                title=slide_title or section.title,
                layout_name=self._select_layout_for_content("image", template_info),
                blocks=slide_blocks
            )
            new_slides.append(slide)
        
        # Add any remaining unused text blocks
        remaining_text_blocks = [block for block in text_blocks if block not in used_blocks]
        if remaining_text_blocks:
            # Group them in slides of reasonable size
            current_blocks = []
            current_size = 0
            current_title = None
            
            for block in remaining_text_blocks:
                # Estimate block size
                block_size = 0
                if block.content.content_type == ContentType.TEXT and block.content.text:
                    block_size = len(block.content.text)
                elif block.content.content_type == ContentType.BULLET_POINTS and block.content.bullet_points:
                    block_size = sum(len(point) for point in block.content.bullet_points)
                
                # Use the first block's title as slide title
                if not current_title and block.title:
                    current_title = block.title
                
                # Start new slide if this block would make it too large
                if current_blocks and current_size + block_size > max_content_per_slide:
                    from uuid import uuid4
                    slide = Slide(
                        id=str(uuid4()),
                        title=current_title or section.title,
                        layout_name=self._select_layout_for_content("text", template_info),
                        blocks=current_blocks.copy()
                    )
                    new_slides.append(slide)
                    current_blocks = []
                    current_size = 0
                    current_title = block.title  # Reset title for new slide
                
                # Add block to current collection
                current_blocks.append(block)
                current_size += block_size
            
            # Add any final blocks
            if current_blocks:
                from uuid import uuid4
                slide = Slide(
                    id=str(uuid4()),
                    title=current_title or section.title,
                    layout_name=self._select_layout_for_content("text", template_info),
                    blocks=current_blocks
                )
                new_slides.append(slide)
        
        # Add other blocks to new slides
        for other_block in other_blocks:
            if other_block in used_blocks:
                continue
                
            from uuid import uuid4
            slide = Slide(
                id=str(uuid4()),
                title=other_block.title or section.title,
                layout_name=self._select_layout_for_content("other", template_info),
                blocks=[other_block]
            )
            new_slides.append(slide)
        
            
        # Log final state
        logger.debug(f"=== Finished content planning for section: '{section.title}' ===")
        logger.debug(f"Created {len(new_slides)} new slides")
        
        # Replace the section's slides
        section.slides = new_slides
        
        return section
    
    def _select_layout_for_content(self, content_type: str, 
                                  template_info: Optional[TemplateInfo]) -> str:
        """
        Select an appropriate layout for the content type.
        
        Args:
            content_type: Type of content (text, table, image, other).
            template_info: Optional template information.
            
        Returns:
            Layout name.
        """
        if not template_info:
            # Default layouts when template_info is not available
            if content_type == "table":
                return "Titre et tableau"
            elif content_type == "image":
                return "Titre et texte 1 visuel gauche"
            elif content_type == "text":
                return "Titre et texte"
            else:
                return "Titre et texte"
        
        # Use template_info to select appropriate layout
        if content_type == "table" and template_info.table_layouts:
            return template_info.table_layouts[0]
        elif content_type == "image" and template_info.image_layouts:
            return template_info.image_layouts[0]
        elif content_type == "text" and template_info.content_layouts:
            return template_info.content_layouts[0]
        elif template_info.layouts:
            # Fall back to first available layout
            return template_info.layouts[0].name
        
        # Ultimate fallback
        return "Titre et texte"