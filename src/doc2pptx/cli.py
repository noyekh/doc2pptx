# src/doc2pptx/cli.py
from __future__ import annotations

import json
import logging
import copy
from importlib.metadata import PackageNotFoundError, version as pkg_version
from pathlib import Path
from typing import Optional, Union, Dict, List
from openai import APIError, AuthenticationError  

import typer
from rich.console import Console
from rich.logging import RichHandler
from rich.progress import Progress

# Import des classes nécessaires
from doc2pptx.core.models import Presentation 
from doc2pptx.llm.optimizer import PresentationOptimizer 
from doc2pptx.core.models import SectionType, SlideBlock, SlideContent, ContentType
from doc2pptx.ppt.template_loader import TemplateLoader

# Configure root logger and doc2pptx logger early with a basic config.
# This ensures that any loggers retrieved at module level (like in settings.py)
# have some handler and level configured from the start.
# The main RichHandler config will replace this later.
logging.basicConfig(level=logging.DEBUG, format="%(name)s - %(levelname)s - %(message)s")

# Import the global settings object here, at the module level
# It should be available to all functions in this module
from doc2pptx.core.settings import settings

# Get the loggers for this module and key sub-modules
logger = logging.getLogger(__name__) # cli_v2 logger
optimizer_logger = logging.getLogger("doc2pptx.llm.optimizer") # Optimizer logger
settings_logger = logging.getLogger("doc2pptx.core.settings") # Settings logger


console = Console()
app = typer.Typer(add_completion=False, no_args_is_help=True)

def _configure_logging(level: Union[str, int] = logging.INFO) -> None:
    """
    Configures logging with RichHandler.
    Should be called once at the application entry point for commands that use it.
    """
    # Map string level to logging level constant if a string is passed
    if isinstance(level, str):
         level_map = {
             "CRITICAL": logging.CRITICAL,
             "ERROR": logging.ERROR,
             "WARNING": logging.WARNING,
             "INFO": logging.INFO,
             "DEBUG": logging.DEBUG,
         }
         level = level_map.get(level.upper(), logging.INFO)


    # Clear existing handlers from the root logger from any basicConfig calls
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    # Set up RichHandler
    logging.basicConfig(
        level=level, # Set the overall level for the root logger
        format="%(message)s", # Rich handles formatting
        handlers=[RichHandler(rich_tracebacks=True, markup=True)],
    )
    # Ensure doc2pptx loggers also use this level, or lower if they are more verbose by default
    logging.getLogger("doc2pptx").setLevel(level) # Set level for the main package logger

    # Set specific levels for potentially noisy libraries
    # Only lower them to DEBUG if the requested level is DEBUG
    if level > logging.DEBUG: # If not in DEBUG mode, set external libraries to WARNING
         logging.getLogger("httpx").setLevel(logging.WARNING)
         logging.getLogger("openai").setLevel(logging.WARNING)
         # Set our own module loggers to the requested level
         logger.setLevel(level)
         optimizer_logger.setLevel(level)
         settings_logger.setLevel(level)
    else: # If in DEBUG mode (-v), set external libraries to DEBUG as well
         logging.getLogger("httpx").setLevel(logging.DEBUG)
         logging.getLogger("openai").setLevel(logging.DEBUG)
         # Set our own module loggers to DEBUG
         logger.setLevel(logging.DEBUG)
         optimizer_logger.setLevel(logging.DEBUG)
         settings_logger.setLevel(logging.DEBUG)


def _default_output_for(input_path: Path) -> Path:
    input_path = Path(input_path)
    try:
        # Calcule le chemin relatif sous data/input
        rel = input_path.relative_to(Path("data") / "input")
    except ValueError:
        # Si ce n’est pas sous data/input, on ne conserve que le nom de fichier
        rel = input_path.name
    # Construit le chemin de sortie sous data/output
    output = Path("data") / "output" / rel
    # Change l’extension en .pptx et résout le chemin absolu
    return output.with_suffix(".pptx").resolve()

def _build_presentation(
    json_path: Path, template_path: Path, output_path: Path, 
    use_ai_optimization: bool = False, use_content_planning: bool = False
) -> Path:
    """
    Loads presentation data, optionally optimizes it using AI, and builds the PPTX.

    Args:
        json_path: Path to the input JSON file.
        template_path: Path to the PowerPoint template file.
        output_path: Path for the output PPTX file.
        use_ai_optimization: Flag to enable AI optimization.
        use_content_planning: Flag to enable AI content planning.

    Returns:
        The Path to the generated PPTX file.

    Raises:
        ValueError: If the input JSON file is invalid.
        RuntimeError: If presentation data validation fails.
        FileNotFoundError: If the input JSON or template file is not found.
        Exception: If an unexpected error occurs during the build process after validation.
    """
    logger.info(f"Loading presentation data from [magenta]{json_path}[/magenta]", extra={"markup": True})
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            presentation_data = json.load(f)
        logger.info("Presentation data loaded successfully.")
    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON file format: {e}")
        raise ValueError(f"Invalid JSON file format: {e}") from e
    except FileNotFoundError:
        logger.error(f"Input JSON file not found: {json_path}")
        raise # Re-raise FileNotFoundError
    except Exception as e:
        logger.error(f"An unexpected error occurred while reading the JSON file: {e}")
        raise # Re-raise other file reading errors

    # Optimization step
    if use_ai_optimization:
        logger.info("AI optimization enabled. Attempting to optimize presentation structure.")
        with Progress(transient=True) as progress:
            task = progress.add_task("[blue]Optimisation IA…", total=100)
            optimizer = None # Initialize optimizer variable
            try:
                progress.update(task, completed=10, description="[blue]Initialisation de l'optimiseur IA…", refresh=True)
                optimizer = PresentationOptimizer()

                if not optimizer.client:
                    progress.update(task, completed=100, description="[red]Optimisation IA désactivée (erreur client)", refresh=True)
                    logger.warning("AI optimization failed: OpenAI client not initialized. Continuing with standard processing.")
                else:
                    progress.update(task, completed=20, description="[blue]Analyse de la présentation par IA…", refresh=True)
                    optimization_result = optimizer.optimize_presentation(presentation_data)
                    progress.update(task, completed=70, description="[blue]Application des recommandations IA…", refresh=True)

                    if optimization_result and isinstance(optimization_result, dict) and "sections" in optimization_result and optimization_result.get("sections"):
                        _apply_optimizations(presentation_data, optimization_result)
                        progress.update(task, completed=100, description="[blue]Optimisation terminée", refresh=True)
                        logger.info("AI optimization recommendations applied.")
                    else:
                        progress.update(task, completed=100, description="[yellow]Optimisation IA a retourné un résultat vide ou invalide.", refresh=True)
                        logger.warning("AI optimization returned empty or invalid result. Continuing without applying optimizations.")

            except Exception as e:
                progress.update(task, completed=100, description=f"[red]Échec de l'optimisation: {e}", refresh=True)
                if not isinstance(e, (APIError, AuthenticationError)):
                    logger.error(f"An unexpected error occurred during AI optimization process: {e}", exc_info=settings.debug)
                logger.warning("AI optimization step failed. Proceeding with presentation build using original data.")
    else:
        logger.info("AI optimization is disabled.")

    # Continue building the presentation
    logger.info("Validating presentation data structure against models...")
    try:
        presentation = Presentation.model_validate(presentation_data)
        logger.info("Presentation data validated successfully.")
    except Exception as e:
        logger.error(f"Failed to validate presentation data against model: {e}")
        raise RuntimeError(f"Invalid presentation data structure: {e}") from e

    logger.info(f"Initializing PPTBuilder with template: [cyan]{template_path}[/cyan]", extra={"markup": True})
    try:
        from doc2pptx.ppt.builder import PPTBuilder
        builder = PPTBuilder(template_path=template_path, use_ai=use_ai_optimization or use_content_planning)
        logger.info("PPTBuilder initialized.")
    except FileNotFoundError:
        logger.error(f"Template file not found: {template_path}")
        raise
    except Exception as e:
        logger.error(f"An error occurred during PPTBuilder initialization: {e}")
        raise

    with Progress(transient=True) as progress:
        task = progress.add_task("[green]Construction…", total=None)
        logger.info("Building presentation...")
        try:
            # Pass content_planning flag to the build method if it accepts it
            if use_content_planning:
                if hasattr(builder, 'use_content_planning'):
                    builder.use_content_planning = True
                else:
                    logger.warning("Content planning requested but not supported by the builder. Continuing without content planning.")
            
            generated = builder.build(presentation, output_path)
            progress.update(task, completed=1, refresh=True)
            logger.info("Presentation built successfully.")
        except Exception as e:
            progress.update(task, completed=1, refresh=True)
            logger.error(f"An error occurred during presentation build: {e}", exc_info=settings.debug)
            raise RuntimeError(f"Presentation build failed: {e}") from e

    return generated


def _apply_optimizations(presentation_data: Dict, optimization_result: Dict) -> None:
    """
    Applies optimization recommendations (layout, type, splits) to the presentation data dictionary.
    """
    logger.info("Applying AI optimization recommendations...")
    sections_map = {section.get("id"): section for section in presentation_data.get("sections", [])}
    
    # Iterate through optimization result sections
    for opt_section in optimization_result.get("sections", []):
        section_id = opt_section.get("id")
        if not section_id:
            logger.warning("Optimization result contained a section without an ID. Skipping.")
            continue

        # Find the corresponding original section
        original_section = sections_map.get(section_id)
        if not original_section:
            logger.warning(f"Optimization result refers to unknown section ID: {section_id}. Skipping section.")
            continue

        # Apply recommended section type
        recommended_type = opt_section.get("recommended_type")
        if recommended_type:
            current_type = original_section.get("type")
            if recommended_type != current_type:
                logger.info(f"Applying recommended section type '{recommended_type}' for section '{section_id}' (was '{current_type}')")
                original_section["type"] = recommended_type

        # Process slides within the section
        updated_slides_list = []
        original_slides_list = original_section.get("slides", [])
        slides_map = {slide.get("id"): slide for slide in original_slides_list}
        
        # Track which original slides have been processed
        processed_slide_ids = set()
        
        # Process optimization recommendations for slides
        for opt_slide in opt_section.get("slides", []):
            slide_id = opt_slide.get("id")
            if not slide_id:
                logger.warning(f"Optimization result contained a slide without an ID in section '{section_id}'. Skipping.")
                continue

            # Find the corresponding original slide
            original_slide = slides_map.get(slide_id)
            if not original_slide:
                logger.warning(f"Optimization result refers to unknown slide ID: {slide_id} in section '{section_id}'. Skipping slide recommendations.")
                continue
                
            # Mark this slide as processed
            processed_slide_ids.add(slide_id)

            # Apply recommended layout
            recommended_layout = opt_slide.get("recommended_layout")
            if recommended_layout:
                current_layout = original_slide.get("layout_name")
                if recommended_layout != current_layout:
                    logger.info(f"Applying recommended layout '{recommended_layout}' for slide '{slide_id}' (was '{current_layout}') in section '{section_id}'.")
                    original_slide["layout_name"] = recommended_layout

            # Handle overflow/split
            overflow = opt_slide.get("overflow_analysis", {})
            split_contents = overflow.get("split_recommendation", [])

            if overflow.get("may_overflow") and len(split_contents) > 1:
                logger.info(f"Overflow indicated for slide {slide_id} in section '{section_id}'. Splitting content into {len(split_contents)} parts.")
                
                # Update original slide with first part and add to updated list
                _update_slide_content(original_slide, split_contents[0])
                updated_slides_list.append(original_slide)
                
                # Create new slides for subsequent parts
                for i, content in enumerate(split_contents[1:], 1):
                    new_slide_id = f"{slide_id}-part-{i+1}"
                    new_slide_title = f"{original_slide.get('title', 'Slide')} (suite {i})"
                    
                    # Create deep copy of original slide
                    new_slide = copy.deepcopy(original_slide)
                    new_slide["id"] = new_slide_id
                    new_slide["title"] = new_slide_title
                    
                    # Update the content
                    _update_slide_content(new_slide, content)
                    
                    # Add to the list
                    updated_slides_list.append(new_slide)
                    logger.info(f"Created new slide '{new_slide_id}' for split content part {i+1}")
            else:
                # No overflow or not enough parts to split, keep original slide
                updated_slides_list.append(original_slide)
        
        # Add any slides that weren't mentioned in the optimization result
        for original_slide in original_slides_list:
            slide_id = original_slide.get("id")
            if slide_id and slide_id not in processed_slide_ids:
                updated_slides_list.append(original_slide)
                logger.debug(f"Preserving original slide '{slide_id}' which wasn't mentioned in optimization result")
        
        # Replace the original slides list with the updated one
        original_section["slides"] = updated_slides_list

    logger.info("Finished applying AI optimization recommendations.")



@app.command("generate")
def generate(
    input_file: Path = typer.Argument(..., exists=True, readable=True, resolve_path=True, help="Path to the input file (JSON or Markdown)."),
    template: Path = typer.Option(..., "--template", "-t", exists=True, readable=True, resolve_path=True, help="Path to the PowerPoint template (.pptx) file."),
    output: Optional[Path] = typer.Option(None, "--output", "-o", writable=True, resolve_path=True, help="Path for the output PPTX file. Defaults to <input_name>.pptx."),
    ai_optimize: bool = typer.Option(False, "--ai-optimize", "-a", help="Enable AI-powered optimization of presentation structure and content layout."),
    content_planning: bool = typer.Option(False, "--content-planning", "-c", help="Enable AI-powered content planning for optimal slide organization."),
    log_level: str = typer.Option("INFO", "--log-level", "-l", show_default=False, help="Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL). Overridden by -v."),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Enable debug logging."),
) -> None:
    """Generates a PowerPoint presentation from a JSON or Markdown input and template."""
    # Determine the effective logging level based on options
    effective_level_str = "DEBUG" if verbose else log_level
    _configure_logging(effective_level_str)

    output = output or _default_output_for(input_file)
    output.parent.mkdir(parents=True, exist_ok=True)
    console.rule("[bold cyan]doc2pptx • generate")
    logger.info(f"Template   : [cyan]{template}[/cyan]", extra={"markup": True})
    logger.info(f"Entrée     : [magenta]{input_file}[/magenta]", extra={"markup": True})
    logger.info(f"Destination: [green]{output}[/green]", extra={"markup": True})

    if ai_optimize:
        logger.info("[bold blue]Optimisation IA activée[/bold blue]", extra={"markup": True})
    else:
        logger.info("[blue]Optimisation IA désactivée[/blue]", extra={"markup": True})

    if content_planning:
        logger.info("[bold blue]Planification de contenu IA activée[/bold blue]", extra={"markup": True})
    else:
        logger.info("[blue]Planification de contenu IA désactivée[/blue]", extra={"markup": True})

    try:
        # Determine file type and process accordingly
        generated_path = _build_presentation_from_file(input_file, template, output, ai_optimize, content_planning)
        console.print(f"\n[bold green]Deck généré :[/] {generated_path}")
    except (ValueError, RuntimeError, FileNotFoundError) as exc:
        console.print(f"[bold red]Erreur de traitement :[/] {exc}")
        raise typer.Exit(code=1) from exc
    except Exception as exc:
        console.print(f"[bold red]Échec génération inattendu :[/] {exc}")
        raise typer.Exit(code=2) from exc

def _build_presentation_from_file(
    input_file: Path, template_path: Path, output_path: Path, 
    use_ai_optimization: bool = False, use_content_planning: bool = False
) -> Path:
    """
    Build a presentation from an input file (JSON or Markdown).
    
    Args:
        input_file: Path to the input file
        template_path: Path to the PowerPoint template
        output_path: Path for the output PPTX
        use_ai_optimization: Flag to enable AI optimization
        use_content_planning: Flag to enable AI content planning
        
    Returns:
        Path to the generated PPTX file
        
    Raises:
        ValueError: If the input file is invalid
        RuntimeError: If data validation fails
        FileNotFoundError: If input file is not found
    """
    # Determine file type based on extension
    file_ext = input_file.suffix.lower()
    
    if file_ext == ".json":
        # Process JSON input
        return _build_presentation(input_file, template_path, output_path, 
                                  use_ai_optimization, use_content_planning)
    
    elif file_ext in (".md", ".markdown"):
        # Process Markdown input
        logger.info(f"Loading presentation data from Markdown: [magenta]{input_file}[/magenta]", 
                    extra={"markup": True})
        try:
            from doc2pptx.ingest.markdown_loader import load_presentation_from_markdown
            
            # Load the presentation from Markdown
            presentation = load_presentation_from_markdown(input_file)
            logger.info("Markdown presentation data loaded successfully.")
            
            # Apply AI optimization if enabled
            if use_ai_optimization:
                logger.info("AI optimization enabled. Analyzing markdown structure.")
                with Progress(transient=True) as progress:
                    task = progress.add_task("[blue]AI Optimization…", total=100)
                    
                    try:
                        # Initialize optimizer
                        progress.update(task, completed=10, description="[blue]Initializing AI optimizer…", refresh=True)
                        optimizer = PresentationOptimizer()
                        
                        if not optimizer.client:
                            progress.update(task, completed=100, description="[red]AI optimization disabled (client error)", refresh=True)
                            logger.warning("AI optimization failed: OpenAI client not initialized.")
                        else:
                            # Optimize presentation structure
                            progress.update(task, completed=30, description="[blue]Analyzing presentation…", refresh=True)
                            presentation_data = presentation.model_dump()
                            optimization_result = optimizer.optimize_presentation(presentation_data)
                            
                            # Apply optimizations
                            progress.update(task, completed=70, description="[blue]Applying AI recommendations…", refresh=True)
                            if optimization_result and isinstance(optimization_result, dict) and "sections" in optimization_result:
                                # Apply optimizations to presentation model
                                _apply_optimizations_to_model(presentation, optimization_result)
                                progress.update(task, completed=100, description="[green]Optimization complete", refresh=True)
                                logger.info("AI optimization recommendations applied.")
                            else:
                                progress.update(task, completed=100, description="[yellow]Optimization returned invalid result", refresh=True)
                                logger.warning("AI optimization returned empty or invalid result.")
                    except Exception as e:
                        progress.update(task, completed=100, description=f"[red]Optimization failed: {e}", refresh=True)
                        logger.error(f"Error during AI optimization: {e}", exc_info=settings.debug)
            
            # Apply content planning if enabled
            if use_content_planning:
                logger.info("Content planning enabled. Optimizing slide organization.")
                with Progress(transient=True) as progress:
                    task = progress.add_task("[blue]Content Planning…", total=100)
                    
                    try:
                        # Initialize content planner
                        progress.update(task, completed=10, description="[blue]Initializing content planner…", refresh=True)
                        from doc2pptx.llm.content_planner import ContentPlanner
                        from doc2pptx.ppt.template_loader import TemplateLoader, TemplateInfo
                        
                        # Create a template loader to analyze template
                        template_loader = TemplateLoader()
                        template_info = template_loader.analyze_template(template_path)
                        
                        planner = ContentPlanner()
                        if not planner.use_ai:
                            progress.update(task, completed=100, description="[yellow]Using basic content planning (AI unavailable)", refresh=True)
                            logger.warning("AI content planning unavailable. Using heuristics.")
                        
                        # Plan content for each section
                        progress.update(task, completed=30, description="[blue]Planning content…", refresh=True)
                        for i, section in enumerate(presentation.sections):
                            try:
                                # Plan section content
                                optimized_section = planner.plan_section_content(section, template_info)
                                presentation.sections[i] = optimized_section
                            except Exception as section_e:
                                logger.warning(f"Error planning content for section '{section.title}': {section_e}")
                        
                        progress.update(task, completed=100, description="[green]Content planning complete", refresh=True)
                        logger.info("Content planning completed successfully.")
                    except Exception as e:
                        progress.update(task, completed=100, description=f"[red]Content planning failed: {e}", refresh=True)
                        logger.error(f"Error during content planning: {e}", exc_info=settings.debug)
            
            # Continue with presentation building
            logger.info(f"Initializing PPTBuilder with template: [cyan]{template_path}[/cyan]", 
                        extra={"markup": True})
            try:
                from doc2pptx.ppt.builder import PPTBuilder
                builder = PPTBuilder(template_path=template_path, 
                                    use_ai=use_ai_optimization, 
                                    use_content_planning=use_content_planning)
                logger.info("PPTBuilder initialized.")
            except FileNotFoundError:
                logger.error(f"Template file not found: {template_path}")
                raise
            except Exception as e:
                logger.error(f"An error occurred during PPTBuilder initialization: {e}")
                raise
            
            with Progress(transient=True) as progress:
                task = progress.add_task("[green]Building…", total=None)
                logger.info("Building presentation...")
                try:
                    # IMPORTANT: Pass presentation as the first argument and output_path as the second
                    # NOT the other way around
                    generated = builder.build(presentation, output_path)
                    progress.update(task, completed=1, refresh=True)
                    logger.info("Presentation built successfully.")
                    return generated
                except Exception as e:
                    progress.update(task, completed=1, refresh=True)
                    logger.error(f"An error occurred during presentation build: {e}", 
                                exc_info=settings.debug)
                    raise RuntimeError(f"Presentation build failed: {e}") from e
                
        except Exception as e:
            logger.error(f"Error processing Markdown file: {e}")
            raise
    
    else:
        # Unsupported file type
        raise ValueError(f"Unsupported input file type: {file_ext}. Expected .json, .md, or .markdown")


def _apply_optimizations_to_model(presentation: Presentation, optimization_result: Dict[str, Any]) -> None:
    """
    Apply AI optimization recommendations to the presentation model.
    
    Args:
        presentation: Presentation model to update
        optimization_result: Optimization recommendations from AI
    """
    # Map sections by ID for easy lookup
    section_map = {section.id: section for section in presentation.sections}
    
    # Process each section in the optimization result
    for opt_section in optimization_result.get("sections", []):
        section_id = opt_section.get("id")
        if not section_id or section_id not in section_map:
            continue
        
        section = section_map[section_id]
        
        # Update section type if recommended
        if "recommended_type" in opt_section:
            try:
                section.type = SectionType(opt_section["recommended_type"])
            except ValueError:
                logger.warning(f"Invalid recommended section type: {opt_section['recommended_type']}")
        
        # Process slides in the section
        if "slides" in opt_section:
            # Map slides by ID for easy lookup
            slide_map = {slide.id: (i, slide) for i, slide in enumerate(section.slides)}
            
            # Track slides that need to be added
            slides_to_add = []
            
            # Process each slide in the optimization result
            for opt_slide in opt_section["slides"]:
                slide_id = opt_slide.get("id")
                if not slide_id or slide_id not in slide_map:
                    continue
                
                index, slide = slide_map[slide_id]
                
                # Update layout if recommended
                if "recommended_layout" in opt_slide:
                    slide.layout_name = opt_slide["recommended_layout"]
                
                # Handle overflow if present
                if "overflow_analysis" in opt_slide:
                    overflow = opt_slide["overflow_analysis"]
                    if overflow.get("may_overflow") and "split_recommendation" in overflow:
                        split_contents = overflow["split_recommendation"]
                        
                        if len(split_contents) > 1:
                            # Update the original slide with the first part
                            _update_slide_content(slide, split_contents[0])
                            
                            # Create new slides for the remaining parts
                            for i, content in enumerate(split_contents[1:], 1):
                                # Create a new slide based on the original
                                from copy import deepcopy
                                new_slide = deepcopy(slide)
                                new_slide.id = f"{slide_id}-part-{i+1}"
                                new_slide.title = f"{slide.title} (suite {i})"
                                
                                # Update content
                                _update_slide_content(new_slide, content)
                                
                                # Add to list of slides to add
                                slides_to_add.append((index + i, new_slide))
            
            # Add new slides
            for i, (insert_index, new_slide) in enumerate(sorted(slides_to_add, key=lambda x: x[0])):
                section.slides.insert(insert_index + i, new_slide)

def _update_slide_content(slide, content: str) -> None:
    """
    Update the content of the first suitable block found in the slide, or creates a new one if needed.
    """
    # Try to find an existing block to update
    found_suitable_block = False
    
    for i, block in enumerate(slide.blocks):
        block_content = block.content
        if not block_content:
            continue
            
        content_type = block_content.content_type
        
        # Check if content looks like bullet points
        content_looks_like_bullets = any(line.strip().startswith(('-', '*', '•')) for line in content.split('\n') if line.strip())
        
        if content_type == ContentType.TEXT and not content_looks_like_bullets:
            # Update text content
            logger.debug(f"Updating text block in slide '{slide.id}'")
            block_content.text = content
            found_suitable_block = True
            break
            
        elif content_type == ContentType.BULLET_POINTS and content_looks_like_bullets:
            # Convert content to bullet points
            logger.debug(f"Updating bullet points block in slide '{slide.id}'")
            bullet_points = []
            for line in content.split('\n'):
                line = line.strip()
                if line:
                    # Remove bullet markers if present
                    if line.startswith(('-', '*', '•')):
                        line = line[1:].strip()
                    bullet_points.append(line)
            
            block_content.bullet_points = bullet_points
            found_suitable_block = True
            break
    
    # If no suitable block found, create a new one
    if not found_suitable_block:
        from doc2pptx.core.models import SlideBlock, SlideContent, ContentType
        slide_id = slide.id
        
        # Determine if content looks like bullet points
        content_looks_like_bullets = any(line.strip().startswith(('-', '*', '•')) for line in content.split('\n') if line.strip())
        
        if content_looks_like_bullets:
            # Process bullet points
            bullet_points = []
            for line in content.split('\n'):
                line = line.strip()
                if line:
                    # Remove bullet markers if present
                    if line.startswith(('-', '*', '•')):
                        line = line[1:].strip()
                    bullet_points.append(line)
                    
            new_block = SlideBlock(
                id=f"{slide_id}-new-block",
                content=SlideContent(
                    content_type=ContentType.BULLET_POINTS,
                    bullet_points=bullet_points
                )
            )
        else:
            # Create text block
            new_block = SlideBlock(
                id=f"{slide_id}-new-block",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text=content
                )
            )
        
        slide.blocks.append(new_block)
        logger.debug(f"Created new content block in slide '{slide_id}'")


@app.command()
def edit(
    presentation: Path = typer.Option(..., "--presentation", "-p", exists=True, readable=True, resolve_path=True, help="Path to the presentation (.pptx) file to edit."),
    command: str = typer.Option(..., "--command", "-c", help="The edit command string (e.g., 'move slide 5 after 10', 'update slide 3 title \"New Title\"')."),
    log_level: str = typer.Option("INFO", "--log-level", "-l", show_default=False, help="Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL). Overridden by -v."),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Enable debug logging."),
) -> None:
    """Edits an existing PowerPoint presentation using a command string."""
    effective_level_str = "DEBUG" if verbose else log_level
    _configure_logging(effective_level_str)
    console.print("[yellow]La commande [bold]edit[/] n'est pas encore disponible (roadmap Jour 7).[/yellow]", markup=True)
    logger.info("Edit command is not yet implemented.")
    raise typer.Exit(code=1)

@app.command()
def prompt(
    presentation: Path = typer.Option(..., "--presentation", "-p", exists=True, readable=True, resolve_path=True, help="Path to the presentation (.pptx) file to edit."),
    nl: str = typer.Option(..., "--nl", help="Natural language prompt for editing."),
    log_level: str = typer.Option("INFO", "--log-level", "-l", show_default=False, help="Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL). Overridden by -v."),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Enable debug logging."),
) -> None:
    """Edits a PowerPoint presentation using a natural language prompt."""
    effective_level_str = "DEBUG" if verbose else log_level
    _configure_logging(effective_level_str)
    console.print("[yellow]La commande [bold]prompt[/] n'est pas encore disponible (roadmap Jour 7).[/yellow]", markup=True)
    logger.info("Prompt command is not yet implemented.")
    raise typer.Exit(code=1)

@app.command()
def version() -> None:
    """Shows the application version."""
    # Configure logging for the version command as well
    _configure_logging(logging.INFO) # Use default INFO level for version command
    logger.info("Displaying version.")
    try:
        # Need to import the package name from pyproject.toml or define it elsewhere
        # For now, hardcode or get from importlib.metadata
        package_name = "doc2pptx"
        console.print(f"{package_name} : [bold green]{pkg_version(package_name)}[/]", markup=True)
    except PackageNotFoundError:
        console.print("[yellow]Package non installé (mode editable ou tests).[/yellow]", markup=True)
        logger.warning(f"Package '{package_name}' not found. Running in editable or test mode?")


def _entrypoint() -> None:
    """Main application entry point."""
    # Logging configuration is now handled within each command function
    # based on user-provided options (-l, -v).
    # A basic config is done at module level for imports.
    app() # Let Typer handle command dispatch

if __name__ == "__main__":
    _entrypoint()