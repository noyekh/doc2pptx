# src/doc2pptx/llm/optimizer.py
import json
import logging
import re
from typing import Dict, List, Optional, Union, Any

from openai import OpenAI, APIError, AuthenticationError

# settings is imported here, runs settings loading logic upon import
from doc2pptx.core.settings import settings

logger = logging.getLogger(__name__) # Get the logger for this module

class PresentationOptimizer:
    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        """
        Initializes the OpenAI PresentationOptimizer.

        Args:
            api_key: Optional OpenAI API key. Defaults to settings.openai_api_key.
            model: Optional OpenAI model name. Defaults to settings.openai_model.
        """
        # Use the provided key or the one from settings
        self.api_key = api_key or settings.openai_api_key
        self.model = model or settings.openai_model
        self.client: Optional[OpenAI] = None # Explicitly type hint the client

        logger.debug(f"Attempting to initialize OpenAI client with model: {self.model}")

        if not self.api_key:
            logger.warning("No OpenAI API key provided. AI optimization will be disabled.")
            # Client remains None
        else:
            # Only log the format warning if the key is not empty
            if not self._is_valid_api_key_format(self.api_key):
                # Log a warning but still try to initialize the client, as the regex might be outdated
                # Log first few chars for debugging
                masked_key_sample = f"{self.api_key[:5]}..." if len(self.api_key) > 5 else self.api_key
                logger.warning(f"API key format ('{masked_key_sample}') seems unusual. Attempting to initialize OpenAI client anyway.")

            try:
                # Attempt to initialize the client
                self.client = OpenAI(api_key=self.api_key)
                # Note: Client initialization itself doesn't validate the key with OpenAI
                # The key is validated on the first API call.
                logger.info(f"OpenAI client initialized successfully.")
            except AuthenticationError as e:
                logger.error(f"OpenAI Authentication Error during client initialization: {e}")
                self.client = None # Ensure client is None on auth failure
            except APIError as e:
                logger.error(f"OpenAI API Error during client initialization: {e}")
                self.client = None # Ensure client is None on other API errors
            except Exception as e:
                logger.error(f"An unexpected error occurred during OpenAI client initialization: {e}")
                self.client = None # Ensure client is None on any other failure

    def _is_valid_api_key_format(self, api_key: Optional[str]) -> bool:
        """
        Checks if the API key format is one of the expected patterns using fullmatch.
        Note: This is a basic format check, not a validation of the key's authenticity or status.
        """
        if not api_key:
            return False
        # Updated regex to include the 'sk-proj-' prefix and be more flexible on length after prefix
        # Using fullmatch ensures the pattern matches the entire string
        valid_formats = [
            r'^sk-proj-[a-zA-Z0-9]+$', # New sk-proj- format (any length after prefix)
            r'^sk-[a-zA-Z0-9]{48}$', # Older sk- format (exactly 48 chars)
            r'^sk-[a-zA-Z0-9]+$', # More general sk- format (any length after prefix, catches org keys too)
            # If sk-org- needs specific handling, add r'^sk-org-[a-zA-Z0-9]+$'
        ]
        # Use re.fullmatch for stricter checking that the pattern matches the entire string
        for pattern in valid_formats:
            if re.fullmatch(pattern, api_key):
                return True
        return False

    def optimize_presentation(self, presentation_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Sends presentation data to OpenAI for optimization recommendations (layout, overflow).
        Returns recommendations or an empty structure on failure.
        """
        if not self.client:
            # This handles cases where the key was missing, invalid format, or initialization failed
            logger.warning("OpenAI client not available. Skipping optimization.")
            return {"sections": []} # Return empty recommendations on failure

        simplified_data = self._simplify_presentation(presentation_data)
        # Keep the prompt largely the same, potentially refine layout names based on available templates
        # Ensure the layout names listed in the prompt exactly match the expected names in your template mapping.
        prompt = f"""
        As a PowerPoint design expert, analyze this presentation structure and provide optimization recommendations:

        {json.dumps(simplified_data, indent=2)}

        For each section and slide, provide:
        1. A validated section type (from standard types: title, introduction, content, conclusion, appendix, custom, agenda, section_header, bullet_list, chart, text_blocks, image_right, two_column, table, image_left, heat_map, quote, numbered_list, thank_you, code, mermaid)
        2. An appropriate layout name from: "Diapositive de titre", "Introduction", "Titre et texte", "Titre et tableau",
           "Titre et texte 1 visuel gauche", "Titre et texte 1 histogramme", "Titre et 3 colonnes", "Chapitre 1"
        3. Whether the content might overflow and should be split across multiple slides

        Return a JSON structure with the same organization, but with recommendations added for each section and slide.
        Follow this exact format:

        {{
            "sections": [
                {{
                    "id": "[section_id]",
                    "recommended_type": "[valid_section_type]",
                    "slides": [
                        {{
                            "id": "[slide_id]",
                            "recommended_layout": "[layout_name]",
                            "overflow_analysis": {{
                                "may_overflow": true/false,
                                "split_recommendation": [
                                    "First slide content",
                                    "Second slide content (if split needed)"
                                ]
                            }}
                        }}
                    ]
                }}
            ]
        }}
        """
        try:
            logger.info("Sending optimization request to OpenAI API")
            # The actual API call where the key is validated by OpenAI
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are a PowerPoint design expert assistant."},
                    {"role": "user", "content": prompt}
                ],
                temperature=settings.openai_temperature, # Use temperature from settings
                max_tokens=4000,
                response_format={"type": "json_object"}
            )
            try:
                # Access the content safely
                response_text = response.choices[0].message.content.strip() if response and response.choices else None
                if not response_text:
                    logger.warning("OpenAI API returned empty response content.")
                    return {"sections": []}

                optimization_result = json.loads(response_text)
                # Basic validation of the JSON structure
                if not isinstance(optimization_result, dict) or "sections" not in optimization_result:
                    logger.warning("Invalid optimization result structure from API")
                    return {"sections": []}

                logger.info("Successfully received and parsed optimization result")
                return optimization_result

            except (json.JSONDecodeError, IndexError, KeyError, AttributeError) as e:
                logger.error(f"Error parsing API response JSON: {e}")
                return {"sections": []}
        except (AuthenticationError, APIError) as e:
            # Catch specific OpenAI errors and log them
            # The traceback is often included in the RichHandler log at DEBUG level
            logger.error(f"OpenAI API Error during optimization call: {e}")
            return {"sections": []}
        except Exception as e:
            # Catch any other unexpected errors
            logger.error(f"An unexpected error occurred during presentation optimization API call: {e}")
            return {"sections": []}

    def _simplify_presentation(self, presentation_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Simplifies presentation data for sending to the LLM, keeping essential structure
        and summaries of content blocks.
        """
        # Implementation remains the same as in the previous step
        simplified = {
            "title": presentation_data.get("title", "Untitled Presentation"),
            "sections": []
        }
        for section in presentation_data.get("sections", []):
            simplified_section = {
                "id": section.get("id", ""),
                "title": section.get("title", ""),
                "type": section.get("type", ""),
                "slides": []
            }
            for slide in section.get("slides", []):
                simplified_slide = {
                    "id": slide.get("id", ""),
                    "title": slide.get("title", ""),
                    "layout_name": slide.get("layout_name", ""),
                    "content_summary": []
                }
                for block in slide.get("blocks", []):
                    content = block.get("content", {})
                    content_type = content.get("content_type", "")
                    content_summary = {"type": content_type}
                    # Added more specific summaries for different content types
                    if content_type == "text" and "text" in content:
                        text = content["text"]
                        content_summary["sample"] = text[:200] + "..." if len(text) > 200 else text
                        content_summary["length"] = len(text)
                    elif content_type == "bullet_points" and "bullet_points" in content:
                        points = content["bullet_points"]
                        content_summary["count"] = len(points)
                        content_summary["sample_first_3"] = points[:3] if len(points) <= 3 else points[:3] + ["..."]
                    elif content_type == "table" and "table" in content:
                        table = content.get("table", {})
                        content_summary["rows"] = len(table.get("rows", []))
                        content_summary["columns"] = len(table.get("headers", []))
                        if table.get("rows"):
                            # Include header and first row sample
                            content_summary["sample_header"] = table.get("headers", [])
                            content_summary["sample_first_row"] = table["rows"][0]
                    elif content_type == "code" and "code" in content:
                        code = content.get("code", {})
                        content_summary["language"] = code.get("language", "")
                        code_text = code.get("code", "")
                        content_summary["sample"] = code_text[:200] + "..." if len(code_text) > 200 else code_text
                    elif content_type == "mermaid" and "mermaid" in content:
                        mermaid = content.get("mermaid", {})
                        mermaid_code = mermaid.get("code", "")
                        content_summary["sample"] = mermaid_code[:200] + "..." if len(mermaid_code) > 200 else mermaid_code
                    elif content_type == "image":
                        image_info = content.get("image", {})
                        content_summary["source"] = image_info.get("source_type", "unknown")
                        if image_info.get("url"):
                             content_summary["url_sample"] = image_info["url"][:50] + "..." if len(image_info["url"]) > 50 else image_info["url"]

                    else:
                        # Include other simple content attributes for types not explicitly summarized
                        content_summary.update({k: v for k, v in content.items() if k != "content"})
                        if "content" in content and isinstance(content["content"], (str, int, float, bool)):
                             content_summary["content_value"] = content["content"]
                        content_summary["note"] = "Detailed summarization not fully implemented for this type."

                    # Add content summary only if it has more than just a type
                    if len(content_summary) > 1:
                         simplified_slide["content_summary"].append(content_summary)

                simplified_section["slides"].append(simplified_slide)
            simplified["sections"].append(simplified_section)
        return simplified

    def suggest_layout(self, section, slide) -> str:
         """Suggests a default layout based on section/slide type (non-AI fallback or initial)."""
         # This method might still be useful as a default if AI optimization is off
         # or if the AI fails to provide a recommendation.
         default_layouts = {
             "title": "Diapositive de titre",
             "introduction": "Introduction",
             "content": "Titre et texte",
             "conclusion": "Chapitre 1",
             "appendix": "Titre et texte",
             "agenda": "Titre et texte", # Assuming agenda is a list
             "section_header": "Chapitre 1",
             "bullet_list": "Titre et texte", # Can be Titre et texte or Titre et 3 colonnes
             "chart": "Titre et texte 1 histogramme",
             "text_blocks": "Titre et texte", # Can be two_column or three_column layouts
             "image_right": "Titre et texte 1 visuel gauche", # Layout name might be counter-intuitive
             "two_column": "Titre et 3 colonnes", # Assuming 3 colonnes layout can handle 2
             "table": "Titre et tableau",
             "image_left": "Titre et texte 1 visuel gauche",
             "heat_map": "Titre et tableau", # Assuming heat map is like a table
             "quote": "Titre et texte",
             "numbered_list": "Titre et texte", # Similar to bullet_list
             "thank_you": "Chapitre 1", # Or a specific Thank You layout
             "code": "Titre et texte",
             "mermaid": "Titre et texte 1 histogramme", # Assuming diagram fits chart placeholder
             "custom": "Titre et texte",
         }
         section_type_str = str(getattr(section, "type", "custom")).lower() # Get type attribute safely
         slide_layout_name = str(getattr(slide, "layout_name", "")).lower() if slide else "" # Get layout name attribute safely

         # If a layout is already specified on the slide, use that first
         if slide_layout_name:
             # Add logic here to validate the existing layout name if needed
             # For simplicity, just return it if present
             return slide_layout_name

         # Otherwise, suggest based on section type
         for type_key, layout_name in default_layouts.items():
             if type_key in section_type_str:
                 return layout_name

         return default_layouts["custom"] # Fallback if no type matches

    def analyze_content_overflow(self, slide, width_pt, height_pt) -> Dict[str, Any]:
        """
        Analyzes content for potential overflow based on text length heuristic (non-AI fallback).
        This is a very rough estimate.
        """
        # This method is likely less accurate than the AI's overflow analysis.
        # Keep it as a fallback if needed, but rely on the AI's result for '--ai-optimize'.
        # The logic here assumes 'slide' is a Pydantic model or object with a 'blocks' attribute.
        if not hasattr(slide, 'blocks') or not slide.blocks:
            return {"may_overflow": False, "split_recommendation": []} # Match AI output keys

        total_text_length = 0
        # Consider different block types and estimate their space usage
        for block in slide.blocks:
            if hasattr(block, 'content') and block.content:
                if hasattr(block.content, 'content_type'):
                    if block.content.content_type == "text" and hasattr(block.content, 'text') and block.content.text:
                        total_text_length += len(block.content.text) * 0.8 # Text length as primary factor
                    elif block.content.content_type == "bullet_points" and hasattr(block.content, 'bullet_points') and block.content.bullet_points:
                        # Estimate space per bullet point (e.g., 50 chars per point)
                        total_text_length += sum(len(str(point)) for point in block.content.bullet_points) * 0.5 + len(block.content.bullet_points) * 50
                    elif block.content.content_type == "table" and hasattr(block.content, 'table') and block.content.table:
                         table = block.content.table
                         # Very rough table size estimate
                         total_text_length += len(table.get("rows", [])) * len(table.get("headers", [])) * 20 # Rows * Cols * estimate_chars_per_cell
                    elif block.content.content_type == "code" and hasattr(block.content, 'code') and block.content.code:
                         code = block.content.code.get("code", "")
                         total_text_length += len(code) * 0.6 # Code might take less space per char

        # Arbitrary threshold based on total text length / estimated space
        # This would ideally be more sophisticated, considering layout, font, etc.
        overflow_threshold = 1500 # Adjusted heuristic threshold

        if total_text_length > overflow_threshold:
            logger.debug(f"Overflow heuristic triggered for slide based on estimated content length: {total_text_length}")
            # Simple split logic (find first text block and split roughly)
            text_to_split = ""
            for block in slide.blocks:
                if hasattr(block, 'content') and block.content and hasattr(block.content, 'content_type') and block.content.content_type == "text" and hasattr(block.content, 'text') and block.content.text:
                    text_to_split = block.content.text
                    break # Only split the first text block in this simple model

            if text_to_split:
                midpoint = len(text_to_split) // 2
                # Try to split at a paragraph or sentence boundary near the midpoint
                split_point = midpoint
                paragraph_split = text_to_split.rfind("\n\n", 0, midpoint)
                sentence_split = max(
                    text_to_split.rfind(". ", 0, midpoint),
                    text_to_split.rfind("! ", 0, midpoint),
                    text_to_split.rfind("? ", 0, midpoint)
                )

                if paragraph_split != -1 and (midpoint - paragraph_split) < (midpoint - sentence_split if sentence_split != -1 else midpoint):
                    split_point = paragraph_split + 2 # Include the newlines
                elif sentence_split != -1:
                    split_point = sentence_split + 1 # Include the punctuation and space

                # Ensure split_point is not 0 or end of string
                if split_point <= 0 or split_point >= len(text_to_split):
                     split_point = midpoint # Fallback to strict midpoint if boundary finding fails

                part1 = text_to_split[:split_point].strip()
                part2 = text_to_split[split_point:].strip()

                if part1 and part2: # Only recommend split if both parts have content
                    logger.debug(f"Overflow split recommended. Split point: {split_point}")
                    return {
                        "may_overflow": True,
                        "split_recommendation": [part1, part2]
                    }
                else:
                     logger.debug("Overflow heuristic triggered, but split resulted in empty part. Skipping split.")
                     return {"may_overflow": False, "split_recommendation": []}
            else:
                logger.debug("Overflow heuristic triggered, but no text block found to split.")
                return {"may_overflow": False, "split_recommendation": []}

        logger.debug("Overflow heuristic did not trigger.")
        return {"may_overflow": False, "split_recommendation": []}


    def validate_and_map_section_type(self, section_type) -> str:
        """
        Validates and maps a given section type string to a standard type.
        Assumes standard_types are defined internally or match SectionType enum.
        """
        # This method might be used to process the AI's 'recommended_type' string.
        # It should ideally map to doc2pptx.core.models.SectionType enum values.
        # Let's update it to use the enum directly.
        from doc2pptx.core.models import SectionType # Import enum here to avoid circular dependency on module level

        # Ensure input is treated as string and is lowercased for case-insensitive matching
        section_type_str = str(section_type).lower()

        # Attempt to match directly to enum values
        try:
             return SectionType(section_type_str).value # Return the standard string value of the enum
        except ValueError:
             # If direct match fails, try mappings
             pass # Continue to mappings

        # Mapping common variations or synonyms to standard types (using enum values)
        type_mapping = {
            "header": SectionType.SECTION_HEADER.value,
            "intro": SectionType.INTRODUCTION.value,
            "summary": SectionType.CONCLUSION.value,
            "list": SectionType.BULLET_LIST.value,
            "graph": SectionType.CHART.value,
            "figure": SectionType.IMAGE_RIGHT.value, # Assuming image_right is default for figures
            "image": SectionType.IMAGE_LEFT.value, # Assuming image_left is default for images
            "text": SectionType.CONTENT.value, # Generic text often maps to content
            "bullets": SectionType.BULLET_LIST.value,
            "columns": SectionType.TWO_COLUMN.value, # Assuming 'columns' implies two columns
            "split": SectionType.TWO_COLUMN.value, # If 'split' means two columns layout
            # Add more mappings based on expected LLM output or common types
        }
        # Check if the input string is a key in the mapping
        if section_type_str in type_mapping:
            logger.debug(f"Mapping section type '{section_type_str}' to standard type '{type_mapping[section_type_str]}'")
            return type_mapping[section_type_str]

        # Check if the input string contains a keyword from the mapping (less precise)
        # Iterate through mapping keys and check if they are substrings
        for key, value in type_mapping.items():
             if key in section_type_str and section_type_str != key: # Avoid re-mapping direct matches
                 logger.debug(f"Mapping section type substring '{key}' in '{section_type_str}' to '{value}'")
                 return value

        logger.warning(f"Unknown section type '{section_type}' - treating as 'custom'")
        return SectionType.CUSTOM.value # Default to the enum value for custom
    
            
    def analyze_template_layouts(self, template_info: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
        """
        Analyze template layouts and provide insights on their optimal use.
        
        Args:
            template_info: Dictionary with template layout information
                
        Returns:
            Dictionary mapping layout names to analysis results
        """
        if not self.client:
            logger.warning("OpenAI client not available. Skipping layout analysis.")
            return {}
        
        # Prepare layout information for analysis
        layout_data = {}
        for layout_name, layout_info in template_info.items():
            layout_data[layout_name] = {
                "has_title": layout_info.get("supports_title", False),
                "has_content": layout_info.get("supports_content", False),
                "has_image": layout_info.get("supports_image", False),
                "has_chart": layout_info.get("supports_chart", False),
                "has_table": layout_info.get("supports_table", False),
                "max_blocks": layout_info.get("max_content_blocks", 0),
                "placeholder_types": layout_info.get("placeholder_types", [])
            }
        
        prompt = f"""
        As a PowerPoint design expert, analyze these slide layouts and provide insights on their optimal use.
        
        For each layout, provide:
        1. A brief description (1-2 sentences)
        2. Best content types to use with this layout (text, bullets, tables, images, charts)
        3. Optimal use cases (e.g., section intro, data presentation, conclusion)
        4. Any limitations to be aware of
        5. A recommendation score (1-10) on how versatile and useful this layout is
        
        Here are the layouts and their capabilities:
        {json.dumps(layout_data, indent=2)}
        
        Return a JSON with this structure:
        {{
            "layout_name1": {{
                "description": "Brief description of the layout",
                "ideal_content_types": ["text", "bullet_points", ...],
                "best_used_for": ["use case 1", "use case 2", ...],
                "limitations": "Any limitations of this layout",
                "recommendation_score": 7
            }},
            ...
        }}
        """
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are a PowerPoint design expert assistant."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                response_format={"type": "json_object"}
            )
            
            response_text = response.choices[0].message.content.strip()
            try:
                analysis_results = json.loads(response_text)
                return analysis_results
            except json.JSONDecodeError:
                logger.error("Failed to parse JSON response from OpenAI")
                return {}
                
        except Exception as e:
            logger.error(f"Error analyzing template layouts: {e}")
            return {}
    
    def analyze_section_content(self, section_content: Dict[str, Any], 
                           available_layouts: Dict[str, Any]) -> Dict[str, Any]:
        """
        Analyze section content and suggest optimal slides organization.
        
        Args:
            section_content: Dictionary containing section content
            available_layouts: Dictionary containing available layouts information
            
        Returns:
            Dictionary with content planning recommendations
        """
        if not self.client:
            logger.warning("OpenAI client not available. Skipping content analysis.")
            return {"slides": []}
        
        prompt = f"""
        As a PowerPoint design expert, suggest the optimal organization of this content across slides.
        
        SECTION CONTENT:
        {json.dumps(section_content, indent=2)}
        
        AVAILABLE LAYOUTS:
        {json.dumps(available_layouts, indent=2)}
        
        Based on the content and available layouts:
        1. Determine how many slides would be optimal
        2. Suggest appropriate layouts for each slide
        3. Indicate how to distribute the content across slides
        4. Consider logical grouping, readability, and visual appeal
        5. For text-heavy content, suggest bullet point conversion where appropriate
        
        Return a detailed content plan in this JSON structure:
        {{
            "slides": [
                {{
                    "title": "Recommended slide title",
                    "layout": "Recommended layout name",
                    "content": [
                        {{
                            "type": "text|bullet_points|table|image",
                            "content": "Actual content to place here",
                            "notes": "Optional explanation of placement"
                        }}
                    ],
                    "reasoning": "Why this content arrangement works well"
                }}
            ],
            "overall_recommendations": "Any general suggestions for improving the content"
        }}
        """
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are a PowerPoint design expert assistant."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                response_format={"type": "json_object"}
            )
            
            response_text = response.choices[0].message.content.strip()
            content_plan = json.loads(response_text)
            return content_plan
            
        except Exception as e:
            logger.error(f"Error analyzing section content: {e}")
            return {"slides": []}