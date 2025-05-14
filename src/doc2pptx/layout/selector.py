"""
Layout selector for doc2pptx.

This module provides functionality to select the appropriate PowerPoint layout
for each section and slide based on rules defined in rules.yaml.
"""
import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union, Any

import yaml
from pptx import Presentation as PptxPresentation

from doc2pptx.core.models import Section, Slide, SectionType, ContentType
from doc2pptx.llm.optimizer import PresentationOptimizer
from doc2pptx.ppt.template_loader import TemplateLoader
import logging

logger = logging.getLogger(__name__)



class LayoutSelector:
    """
    Selects the appropriate PowerPoint layout for each section and slide.
    
    This class applies rules defined in rules.yaml to match section/slide content
    with the most appropriate layout in the PowerPoint template.
    """
        
    def __init__(self, rules_path: Optional[Union[str, Path]] = None, template: Optional[PptxPresentation] = None, use_ai: bool = True):
        """
        Initialize a layout selector with rules from a YAML file.
        
        Args:
            rules_path: Path to the YAML file containing layout selection rules.
                        Defaults to the rules.yaml in the same directory.
            template: Optional PowerPoint presentation template to get available layouts.
                        If provided, validates that all layouts in rules exist in template.
            use_ai: Whether to use AI optimization for layout selection.
        
        Raises:
            FileNotFoundError: If the rules file does not exist.
            ValueError: If the rules file is invalid or layouts don't exist in template.
        """
        if rules_path is None:
            # Use default rules.yaml in the same directory as this file
            current_dir = Path(__file__).parent
            rules_path = current_dir / "rules.yaml"
        
        self.rules_path = Path(rules_path)
        
        if not self.rules_path.exists():
            raise FileNotFoundError(f"Rules file not found: {self.rules_path}")
        
        # Load rules from YAML file
        self.rules = self._load_rules()
        
        # Store available layouts from template if provided
        self.available_layouts = []
        if template is not None:
            self.available_layouts = [layout.name for layout in template.slide_layouts]
            self._validate_rules_against_template()
            
            # Get template info using TemplateLoader
            template_loader = TemplateLoader()
            if use_ai:
                self.template_info = template_loader.analyze_template_with_ai(template)
            else:
                self.template_info = template_loader.analyze_template(template)
        else:
            self.template_info = None
                
        # Initialize AI optimizer if enabled
        self.use_ai = use_ai
        if self.use_ai:
            try:
                self.optimizer = PresentationOptimizer()
            except Exception as e:
                logger.warning(f"Could not initialize AI optimizer: {e}. Falling back to rule-based selection.")
                self.use_ai = False
        
    def get_layout_name(self, section: Section, slide: Optional[Slide] = None) -> str:
        """
        Get the appropriate layout name for a section or slide.
        """
        # If slide already has a layout_name and it's not 'auto', use it
        if slide and slide.layout_name and slide.layout_name != "auto":
            # Check if this layout exists in the template if available_layouts is populated
            if self.available_layouts and slide.layout_name not in self.available_layouts:
                # If layout doesn't exist, we'll select an appropriate one
                pass
            else:
                # MODIFICATION: Vérifier si le slide contient une table, même si un layout est déjà défini
                has_table = any(block.content and block.content.content_type == ContentType.TABLE for block in slide.blocks)
                if has_table and slide.layout_name != "Titre et tableau" and (not self.template_info or 
                                                                            slide.layout_name not in self.template_info.table_layouts):
                    # On a une table mais pas un layout de table, on ignore le layout spécifié
                    logger.warning(f"Slide with title '{slide.title if slide.title else '(No title)'}' contains a table but "
                                f"has layout '{slide.layout_name}' which may not support tables. "
                                f"Selecting an appropriate table layout instead.")
                    # Continue to table layout selection below
                else:
                    # Layout is appropriate, use it
                    return slide.layout_name
        
        # MODIFICATION: Vérifier d'abord si le slide contient une table
        # Cette vérification doit être prioritaire sur toutes les autres
        if slide:
            has_table = any(block.content and block.content.content_type == ContentType.TABLE for block in slide.blocks)
            if has_table:
                # Sélectionner un layout de table
                table_layout = self.rules.get("table_layout", "Titre et tableau")
                if not self.available_layouts or table_layout in self.available_layouts:
                    logger.info(f"Selected table layout '{table_layout}' for slide containing a table")
                    return table_layout
                elif self.available_layouts:
                    # Si le layout de table par défaut n'existe pas, chercher un autre layout pour les tables
                    if hasattr(self, 'template_info') and self.template_info and self.template_info.table_layouts:
                        logger.info(f"Selected first table layout '{self.template_info.table_layouts[0]}' for slide containing a table")
                        return self.template_info.table_layouts[0]
                    
                    # Rechercher un layout qui pourrait supporter les tables
                    for layout in self.available_layouts:
                        if "tableau" in layout.lower() or "table" in layout.lower():
                            logger.info(f"Selected alternative table layout '{layout}' for slide containing a table")
                            return layout    
        
        # Try AI-based layout selection first if enabled
        if self.use_ai and hasattr(self, 'optimizer') and self.optimizer and self.optimizer.client:
            try:
                ai_layout = self.optimizer.suggest_layout(section, slide)
                # Validate that the suggested layout exists in the template
                if not self.available_layouts or ai_layout in self.available_layouts:
                    logger.info(f"Using AI-suggested layout: {ai_layout}")
                    return ai_layout
                else:
                    logger.warning(f"AI suggested layout '{ai_layout}' not found in template.")
            except Exception as e:
                logger.warning(f"Error in AI layout selection: {e}")
        
        # Default layout if nothing matches
        default_layout = self.rules.get("default_layout", "Titre et texte")
        
        # If no slide provided, use section-based layout or default
        if not slide:
            section_type_layout = self.rules["section_types"].get(section.type.value, None)
            return section_type_layout or default_layout
        
        # Process content combinations
        if "content_combinations" in self.rules:
            for combo_rule in self.rules["content_combinations"]:
                if self._matches_combination(slide, combo_rule):
                    layout = combo_rule["layout"]
                    if not self.available_layouts or layout in self.available_layouts:
                        return layout
        
        # Check for specific content types
        has_table = any(block.content and block.content.content_type == ContentType.TABLE for block in slide.blocks)
        has_image = any(block.content and block.content.content_type == ContentType.IMAGE for block in slide.blocks)
        has_chart = any(block.content and block.content.content_type in [ContentType.CHART, ContentType.MERMAID] for block in slide.blocks)
        
        # Special case for tables
        if has_table:
            table_layout = self.rules.get("table_layout", "Titre et tableau")
            if not self.available_layouts or table_layout in self.available_layouts:
                return table_layout
        
        # Special case for images
        if has_image:
            image_layout = self.rules.get("image_layout", "Titre et texte 1 visuel gauche")
            if not self.available_layouts or image_layout in self.available_layouts:
                return image_layout
        
        # Special case for charts
        if has_chart:
            chart_layout = self.rules.get("chart_layout", "Titre et texte 1 histogramme")
            if not self.available_layouts or chart_layout in self.available_layouts:
                return chart_layout
        
        # Check number of blocks
        num_blocks = len(slide.blocks)
        
        if num_blocks > 2:
            multi_block_layout = self.rules.get("multi_block_layout", "Titre et 3 colonnes")
            if not self.available_layouts or multi_block_layout in self.available_layouts:
                return multi_block_layout
        elif num_blocks == 2:
            two_block_layout = self.rules.get("two_block_layout", "Titre et texte")
            if not self.available_layouts or two_block_layout in self.available_layouts:
                return two_block_layout
        
        # Try content type
        if slide.blocks and slide.blocks[0].content:
            content_type = slide.blocks[0].content.content_type.value
            content_type_layout = self.rules["content_types"].get(content_type)
            
            if content_type_layout and content_type_layout != "auto":
                if not self.available_layouts or content_type_layout in self.available_layouts:
                    return content_type_layout
        
        # Try section type
        section_type_layout = self.rules["section_types"].get(section.type.value)
        if section_type_layout and section_type_layout != "auto":
            if not self.available_layouts or section_type_layout in self.available_layouts:
                return section_type_layout
        
        # Last resort - use default
        return default_layout

    
    def _select_layout_with_ai_insights(self, section: Section, slide: Optional[Slide], 
                                       default_layout: str) -> str:
        """
        Select a layout using AI-enhanced template insights.
        
        Args:
            section: The section to get a layout for.
            slide: Optional slide to get a more specific layout for.
            default_layout: Default layout to use if no match is found.
            
        Returns:
            The name of the selected layout.
        """
        if not slide:
            # If no slide provided, use section-based layout or default
            section_type_layout = self.rules["section_types"].get(section.type.value, None)
            return section_type_layout or default_layout
        
        # Determine content types in the slide
        content_types = []
        for block in slide.blocks:
            if block.content:
                content_types.append(block.content.content_type.value)
        
        # Find layouts that support these content types
        candidate_layouts = []
        for layout_name, layout_info in self.template_info.layout_map.items():
            # Check if layout supports the content types
            supports_needed_content = True
            
            # Check number of content blocks
            if len(slide.blocks) > layout_info.max_content_blocks and layout_info.max_content_blocks > 0:
                supports_needed_content = False
                continue
            
            # Check content types compatibility
            for content_type in content_types:
                if content_type == "text" and not layout_info.supports_content:
                    supports_needed_content = False
                    break
                elif content_type == "bullet_points" and not layout_info.supports_content:
                    supports_needed_content = False
                    break
                elif content_type == "table" and not layout_info.supports_table:
                    supports_needed_content = False
                    break
                elif content_type == "image" and not layout_info.supports_image:
                    supports_needed_content = False
                    break
                elif content_type == "chart" and not layout_info.supports_chart:
                    supports_needed_content = False
                    break
            
            if supports_needed_content:
                # Use layout recommendation score from AI insights
                score = layout_info.recommendation_score
                
                # Boost score if content types match ideal content types
                for content_type in content_types:
                    if content_type in layout_info.ideal_content_types:
                        score += 2
                
                # Boost score if section type is in best_used_for
                if section.type.value in [use_case.lower() for use_case in layout_info.best_used_for]:
                    score += 2
                
                candidate_layouts.append((layout_name, score))
        
        # Sort candidates by score (higher is better)
        candidate_layouts.sort(key=lambda x: x[1], reverse=True)
        
        if candidate_layouts:
            return candidate_layouts[0][0]
        
        # Fall back to rule-based selection if no AI insights match
        return self._select_layout_with_rules(section, slide, default_layout)
    
    def _select_layout_with_rules(self, section: Section, slide: Optional[Slide], 
                                 default_layout: str) -> str:
        """
        Select a layout using pure rule-based selection.
        This is the original implementation of get_layout_name.
        
        Args:
            section: The section to get a layout for.
            slide: Optional slide to get a more specific layout for.
            default_layout: Default layout to use if no match is found.
            
        Returns:
            The name of the selected layout.
        """
        # Try AI-based layout selection first if enabled
        if self.use_ai and slide:
            try:
                ai_layout = self.optimizer.suggest_layout(section, slide)
                # Validate that the suggested layout exists in the template
                if self.available_layouts and ai_layout in self.available_layouts:
                    logger.info(f"Using AI-suggested layout: {ai_layout}")
                    return ai_layout
                elif not self.available_layouts:
                    # If we don't know available layouts, trust the AI suggestion
                    logger.info(f"Using AI-suggested layout: {ai_layout} (template layouts unknown)")
                    return ai_layout
                else:
                    logger.warning(f"AI suggested layout '{ai_layout}' not found in template. Using rule-based selection.")
            except Exception as e:
                logger.warning(f"Error in AI layout selection: {e}. Falling back to rule-based selection.")
        

    
    def _load_rules(self) -> Dict[str, Any]:
        """
        Load layout selection rules from the YAML file.
        
        Returns:
            Dict containing the rules for layout selection.
            
        Raises:
            ValueError: If the rules file is invalid.
        """
        try:
            with open(self.rules_path, "r", encoding="utf-8") as f:
                rules = yaml.safe_load(f)
            
            # Validate rules structure
            if not isinstance(rules, dict):
                raise ValueError("Rules must be a dictionary")
            
            if "section_types" not in rules:
                raise ValueError("Rules must contain 'section_types' mapping")
            
            if "content_types" not in rules:
                raise ValueError("Rules must contain 'content_types' mapping")
            
            # Ensure all section types are valid
            for section_type in rules["section_types"]:
                try:
                    SectionType(section_type)
                except ValueError:
                    # Skip validation for custom section types that might not be in enum
                    if section_type != "default" and not section_type.startswith("custom_"):
                        raise ValueError(f"Invalid section type in rules: {section_type}")
            
            # Ensure all content types are valid
            for content_type in rules["content_types"]:
                try:
                    ContentType(content_type)
                except ValueError:
                    # Skip validation for custom content types that might not be in enum
                    if content_type != "default" and not content_type.startswith("custom_"):
                        raise ValueError(f"Invalid content type in rules: {content_type}")
            
            return rules
                
        except yaml.YAMLError as e:
            raise ValueError(f"Invalid YAML in rules file: {e}")
    
    def _validate_rules_against_template(self) -> None:
        """
        Validate that all layouts mentioned in rules exist in the template.
        
        Raises:
            ValueError: If layouts in rules don't exist in template.
        """
        if not self.available_layouts:
            return
        
        # Check section type layouts
        for section_type, layout_name in self.rules["section_types"].items():
            if layout_name not in self.available_layouts and layout_name != "auto":
                raise ValueError(f"Layout '{layout_name}' specified for section type '{section_type}' not found in template")
        
        # Check content type layouts
        for content_type, layout_name in self.rules["content_types"].items():
            if layout_name not in self.available_layouts and layout_name != "auto":
                raise ValueError(f"Layout '{layout_name}' specified for content type '{content_type}' not found in template")
        
        # Check content pattern layouts
        if "content_patterns" in self.rules:
            for pattern, layout_name in self.rules["content_patterns"].items():
                if layout_name not in self.available_layouts:
                    raise ValueError(f"Layout '{layout_name}' specified for content pattern '{pattern}' not found in template")
    
    # def get_layout_name(self, section: Section, slide: Optional[Slide] = None) -> str:
    #     """
    #     Get the appropriate layout name for a section or slide.
        
    #     This method applies the rules to select the most appropriate layout name
    #     based on the section type, slide content type, and other factors.
        
    #     Args:
    #         section: The section to get a layout for.
    #         slide: Optional slide to get a more specific layout for.
    #                If not provided, selects based on section info only.
        
    #     Returns:
    #         The name of the selected layout.
    #     """
    #     # If slide already has a layout_name and it's not 'auto', use it
    #     if slide and slide.layout_name and slide.layout_name != "auto":
    #         # Check if this layout exists in the template if available_layouts is populated
    #         if self.available_layouts and slide.layout_name not in self.available_layouts:
    #             # If layout doesn't exist, we'll select an appropriate one
    #             pass
    #         else:
    #             return slide.layout_name
        
    #     # Default layout if nothing matches
    #     default_layout = self.rules.get("default_layout", "Titre et texte")
        
    #     # If no slide provided, use section-based layout or default
    #     if not slide:
    #         section_type_layout = self.rules["section_types"].get(section.type.value, None)
    #         return section_type_layout or default_layout
        
    #     # For debugging
    #     print(f"Processing slide: {slide.title}, with {len(slide.blocks)} blocks")
        
    #     # Now determine the layout based on slide content with the following priorities:
        
    #     # 1. FIRST PRIORITY: Content combinations
    #     if "content_combinations" in self.rules:
    #         for combo_rule in self.rules["content_combinations"]:
    #             if self._matches_combination(slide, combo_rule):
    #                 # print(f"Matched combination rule: {combo_rule}")
    #                 return combo_rule["layout"]
        
    #     # 2. SECOND PRIORITY: Block count
    #     num_blocks = len(slide.blocks)
    #     if num_blocks == 2:
    #         # print("Using two_block_layout")
    #         two_block_layout = self.rules.get("two_block_layout")
    #         if two_block_layout:
    #             return two_block_layout
    #     elif num_blocks > 3:
    #         # print("Using multi_block_layout")
    #         multi_block_layout = self.rules.get("multi_block_layout")
    #         if multi_block_layout:
    #             return multi_block_layout
        
    #     # 3. THIRD PRIORITY: Content patterns (only for text content)
    #     if slide.blocks and slide.blocks[0].content and slide.blocks[0].content.content_type == ContentType.TEXT:
    #         if "content_patterns" in self.rules:
    #             content_text = slide.blocks[0].content.text or ""
    #             # print(f"Checking pattern match for: {content_text}")
    #             for pattern, layout_name in self.rules["content_patterns"].items():
    #                 if re.search(pattern, content_text, re.IGNORECASE):
    #                     # print(f"Matched pattern: {pattern}")
    #                     return layout_name
        
    #     # 4. FOURTH PRIORITY: Primary content type
    #     if slide.blocks and slide.blocks[0].content:
    #         primary_content_type = slide.blocks[0].content.content_type
    #         # print(f"Checking content type: {primary_content_type.value}")
    #         content_type_layout = self.rules["content_types"].get(primary_content_type.value)
    #         if content_type_layout and content_type_layout != "auto":
    #             # print(f"Using content type layout: {content_type_layout}")
    #             return content_type_layout
        
    #     # 5. FIFTH PRIORITY: Section type
    #     section_type_layout = self.rules["section_types"].get(section.type.value)
    #     if section_type_layout and section_type_layout != "auto":
    #         # print(f"Using section type layout: {section_type_layout}")
    #         return section_type_layout
        
    #     # Final fallback
    #     # print(f"Using default layout: {default_layout}")
    #     return default_layout
    
    def _matches_combination(self, slide: Slide, combo_rule: Dict[str, Any]) -> bool:
        """
        Check if a slide matches a content combination rule.
        
        Args:
            slide: The slide to check.
            combo_rule: The combination rule to match against.
        
        Returns:
            True if the slide matches the combination rule, False otherwise.
        """
        if "requires" not in combo_rule:
            return False
        
        requirements = combo_rule["requires"]
        
        # Check if all required content types are present
        if "content_types" in requirements:
            required_types = set(requirements["content_types"])
            
            # Collect all content types from all blocks
            slide_types = set()
            for block in slide.blocks:
                if block.content:
                    slide_types.add(block.content.content_type.value)
            
            # print(f"Required types: {required_types}, Slide types: {slide_types}")
            if not required_types.issubset(slide_types):
                return False
        
        # Check if slide has specific number of blocks
        if "block_count" in requirements:
            block_count = requirements["block_count"]
            if isinstance(block_count, int) and len(slide.blocks) != block_count:
                return False
            elif isinstance(block_count, dict):
                if "min" in block_count and len(slide.blocks) < block_count["min"]:
                    return False
                if "max" in block_count and len(slide.blocks) > block_count["max"]:
                    return False
        
        # Check for specific title patterns
        if "title_pattern" in requirements and slide.title:
            pattern = requirements["title_pattern"]
            # print(f"Checking title pattern: {pattern} against {slide.title}")
            if not re.search(pattern, slide.title, re.IGNORECASE):
                return False
        
        # All requirements met
        return True