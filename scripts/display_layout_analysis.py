#!/usr/bin/env python3
"""
Utilitaire pour analyser les layouts d'un template PowerPoint et
tester le chargement d'un fichier markdown.
"""

import argparse
import json
import logging
from pathlib import Path
from pprint import pprint

from doc2pptx.ppt.template_loader import TemplateLoader
from doc2pptx.ingest.markdown_loader import load_presentation_from_markdown

# Configuration du logging
logging.basicConfig(level=logging.INFO, format="%(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger("doc2pptx-analyzer")


def analyze_template(template_path: Path, use_ai: bool = False, output_file: Path = None):
    """
    Analyse un template PowerPoint et affiche les informations détaillées sur les layouts.
    
    Args:
        template_path: Chemin vers le fichier template .pptx
        use_ai: Utiliser l'IA pour améliorer l'analyse (nécessite une clé API)
        output_file: Fichier de sortie pour sauvegarder les résultats (optionnel)
    """
    logger.info(f"Analyse du template: {template_path}")
    
    loader = TemplateLoader()
    
    try:
        # Analyser le template
        if use_ai:
            logger.info("Utilisation de l'IA pour l'analyse (peut prendre plus de temps)")
            template_info = loader.analyze_template_with_ai(template_path)
        else:
            template_info = loader.analyze_template(template_path)
        
        # Afficher les informations sur les layouts
        print("\n===== LAYOUTS DÉTECTÉS =====")
        for layout_name, layout_info in template_info.layout_map.items():
            print(f"\nLayout: {layout_name}")
            print(f"  Index: {layout_info.idx}")
            print(f"  Supporte titre: {layout_info.supports_title}")
            print(f"  Supporte contenu: {layout_info.supports_content}")
            print(f"  Supporte image: {layout_info.supports_image}")
            print(f"  Supporte graphique: {layout_info.supports_chart}")
            print(f"  Supporte tableau: {layout_info.supports_table}")
            print(f"  Blocs de contenu max: {layout_info.max_content_blocks}")
            print(f"  Placeholder types: {layout_info.placeholder_types}")
            
            if use_ai and hasattr(layout_info, 'ai_description') and layout_info.ai_description:
                print(f"  Description IA: {layout_info.ai_description}")
                print(f"  Utilisations recommandées: {layout_info.best_used_for}")
                print(f"  Types de contenu idéaux: {layout_info.ideal_content_types}")
                print(f"  Limitations: {layout_info.limitations}")
                print(f"  Score de recommandation: {layout_info.recommendation_score}")
        
        # Afficher les listes spéciales
        print("\n===== LISTES SPÉCIALES =====")
        print(f"Layouts de titre: {template_info.title_layouts}")
        print(f"Layouts de contenu: {template_info.content_layouts}")
        print(f"Layouts d'image: {template_info.image_layouts}")
        print(f"Layouts de graphique: {template_info.chart_layouts}")
        print(f"Layouts de tableau: {template_info.table_layouts}")
        print(f"Layouts à deux contenus: {template_info.two_content_layouts}")
        
        # Comparer avec le dictionnaire statique
        print("\n===== COMPARAISON AVEC DICTIONNAIRE STATIQUE =====")
        from doc2pptx.ppt.builder_v4 import LAYOUT_CAPABILITIES
        
        for layout_name, layout_info in template_info.layout_map.items():
            if layout_name in LAYOUT_CAPABILITIES:
                static_info = LAYOUT_CAPABILITIES[layout_name]
                print(f"\nLayout: {layout_name}")
                print(f"  Statique - titre: {static_info.get('title')}, Détecté: {layout_info.supports_title}")
                print(f"  Statique - contenu: {static_info.get('content')}, Détecté: {layout_info.supports_content}")
                print(f"  Statique - image: {static_info.get('image')}, Détecté: {layout_info.supports_image}")
                print(f"  Statique - graphique: {static_info.get('chart')}, Détecté: {layout_info.supports_chart}")
                print(f"  Statique - tableau: {static_info.get('table')}, Détecté: {layout_info.supports_table}")
                print(f"  Statique - blocs max: {static_info.get('max_blocks')}, Détecté: {layout_info.max_content_blocks}")
            else:
                print(f"\nLayout: {layout_name} - Non présent dans le dictionnaire statique")
        
        # Sauvegarder les résultats si demandé
        if output_file:
            output_data = {
                "layouts": {name: {
                    "supports_title": info.supports_title,
                    "supports_content": info.supports_content,
                    "supports_image": info.supports_image,
                    "supports_chart": info.supports_chart,
                    "supports_table": info.supports_table,
                    "max_content_blocks": info.max_content_blocks,
                    "placeholder_types": [str(pt) for pt in info.placeholder_types],
                    "placeholder_names": info.placeholder_names,
                } for name, info in template_info.layout_map.items()},
                "title_layouts": template_info.title_layouts,
                "content_layouts": template_info.content_layouts,
                "image_layouts": template_info.image_layouts,
                "chart_layouts": template_info.chart_layouts,
                "table_layouts": template_info.table_layouts,
                "two_content_layouts": template_info.two_content_layouts,
            }
            
            # Ajouter les informations IA si disponibles
            if use_ai:
                for name, info in template_info.layout_map.items():
                    if hasattr(info, 'ai_description') and info.ai_description:
                        output_data["layouts"][name]["ai_description"] = info.ai_description
                        output_data["layouts"][name]["best_used_for"] = info.best_used_for
                        output_data["layouts"][name]["ideal_content_types"] = info.ideal_content_types
                        output_data["layouts"][name]["limitations"] = info.limitations
                        output_data["layouts"][name]["recommendation_score"] = info.recommendation_score
            
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(output_data, f, indent=2, ensure_ascii=False)
            logger.info(f"Résultats sauvegardés dans {output_file}")
            
    except Exception as e:
        logger.error(f"Erreur lors de l'analyse du template: {e}")
        import traceback
        traceback.print_exc()


def test_markdown_loader(markdown_path: Path, output_file: Path = None):
    """
    Teste le chargement d'un fichier markdown et affiche la structure résultante.
    
    Args:
        markdown_path: Chemin vers le fichier markdown
        output_file: Fichier de sortie pour sauvegarder les résultats (optionnel)
    """
    logger.info(f"Test du chargement du markdown: {markdown_path}")
    
    try:
        # Charger le markdown
        presentation = load_presentation_from_markdown(markdown_path)
        
        # Afficher les informations de base
        print("\n===== PRÉSENTATION CHARGÉE =====")
        print(f"Titre: {presentation.title}")
        print(f"Auteur: {presentation.author}")
        print(f"Description: {presentation.description}")
        
        # Afficher les sections et slides
        print("\n===== SECTIONS ET SLIDES =====")
        for i, section in enumerate(presentation.sections, 1):
            print(f"\nSection {i}: {section.title} (Type: {section.type})")
            
            for j, slide in enumerate(section.slides, 1):
                print(f"  Slide {j}: {slide.title} (Layout: {slide.layout_name})")
                print(f"    Nombre de blocs: {len(slide.blocks)}")
                
                for k, block in enumerate(slide.blocks, 1):
                    if block.content:
                        content_type = block.content.content_type
                        print(f"      Bloc {k}: Type={content_type.value}")
                        
                        if content_type == "text" and block.content.text:
                            text_preview = block.content.text[:50] + "..." if len(block.content.text) > 50 else block.content.text
                            print(f"        Texte: {text_preview}")
                        elif content_type == "bullet_points" and block.content.bullet_points:
                            points_count = len(block.content.bullet_points)
                            print(f"        Points: {points_count}")
                            if points_count > 0:
                                point_preview = block.content.bullet_points[0][:50] + "..." if len(block.content.bullet_points[0]) > 50 else block.content.bullet_points[0]
                                print(f"        Premier point: {point_preview}")
                        elif content_type == "table" and block.content.table:
                            headers = block.content.table.headers
                            rows = block.content.table.rows
                            print(f"        Tableau: {len(headers)} colonnes, {len(rows)} lignes")
                            print(f"        En-têtes: {headers}")
                        elif content_type == "image" and block.content.image:
                            print(f"        Image: {block.content.image}")
        
        # Sauvegarder les résultats si demandé
        if output_file:
            # Convertir en dictionnaire pour JSON
            output_data = presentation.model_dump()
            
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(output_data, f, indent=2, ensure_ascii=False)
            logger.info(f"Résultats sauvegardés dans {output_file}")
            
    except Exception as e:
        logger.error(f"Erreur lors du chargement du markdown: {e}")
        import traceback
        traceback.print_exc()


def main():
    parser = argparse.ArgumentParser(description="Analyseur pour doc2pptx - templates et markdown")
    subparsers = parser.add_subparsers(dest="command", help="Commande à exécuter")
    
    # Sous-commande pour l'analyse de template
    template_parser = subparsers.add_parser("template", help="Analyser un template PowerPoint")
    template_parser.add_argument("template_path", type=Path, help="Chemin vers le fichier template .pptx")
    template_parser.add_argument("--ai", action="store_true", help="Utiliser l'IA pour améliorer l'analyse")
    template_parser.add_argument("--output", "-o", type=Path, help="Fichier de sortie pour les résultats (JSON)")
    
    # Sous-commande pour le test du markdown_loader
    markdown_parser = subparsers.add_parser("markdown", help="Tester le chargement d'un fichier markdown")
    markdown_parser.add_argument("markdown_path", type=Path, help="Chemin vers le fichier markdown")
    markdown_parser.add_argument("--output", "-o", type=Path, help="Fichier de sortie pour les résultats (JSON)")
    
    args = parser.parse_args()
    
    if args.command == "template":
        analyze_template(args.template_path, args.ai, args.output)
    elif args.command == "markdown":
        test_markdown_loader(args.markdown_path, args.output)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()