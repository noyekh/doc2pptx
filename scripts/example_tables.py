"""
Example script demonstrating the improved table handling in doc2pptx.

This script creates a presentation with various types of tables
to showcase the new capabilities of the PPTBuilder.
"""
import sys
from pathlib import Path
import argparse

# Add the project root to the Python path
project_root = Path(__name__).parent
sys.path.append(str(project_root))

from doc2pptx.core.models import (
    Presentation, Section, Slide, SlideBlock, SlideContent, 
    ContentType, TableData, SectionType
)
from doc2pptx.ppt.builder_v3 import PPTBuilder


def create_simple_table_slide():
    """Create a slide with a simple table."""
    return Slide(
        id="simple_table",
        title="Table Basique",
        layout_name="Titre et tableau",
        blocks=[
            SlideBlock(
                id="simple_table_block",
                content=SlideContent(
                    content_type=ContentType.TABLE,
                    table=TableData(
                        headers=["Produit", "Prix", "Quantité", "Total"],
                        rows=[
                            ["Ordinateur portable", "999 €", "2", "1998 €"],
                            ["Écran 27\"", "299 €", "3", "897 €"],
                            ["Clavier", "79 €", "5", "395 €"],
                            ["Souris", "45 €", "5", "225 €"],
                            ["Casque audio", "125 €", "2", "250 €"]
                        ]
                    )
                )
            )
        ]
    )


def create_styled_table_slide():
    """Create a slide with a styled table."""
    return Slide(
        id="styled_table",
        title="Table avec Style",
        layout_name="Titre et tableau",
        blocks=[
            SlideBlock(
                id="styled_table_block",
                content=SlideContent(
                    content_type=ContentType.TABLE,
                    table=TableData(
                        headers=["Trimestre", "Chiffre d'affaires", "Charges", "Résultat", "style:accent1"],
                        rows=[
                            ["T1 2025", "1,245,000 €", "980,000 €", "265,000 €"],
                            ["T2 2025", "1,450,000 €", "1,050,000 €", "400,000 €"],
                            ["T3 2025", "1,320,000 €", "995,000 €", "325,000 €"],
                            ["T4 2025", "1,680,000 €", "1,180,000 €", "500,000 €"],
                            ["Total", "5,695,000 €", "4,205,000 €", "1,490,000 €"]
                        ]
                    )
                )
            )
        ]
    )


def create_formatted_table_slide():
    """Create a slide with formatted cell content."""
    return Slide(
        id="formatted_table",
        title="Contenu de Table Formaté",
        layout_name="Titre et tableau",
        blocks=[
            SlideBlock(
                id="formatted_table_block",
                content=SlideContent(
                    content_type=ContentType.TABLE,
                    table=TableData(
                        headers=["Projet", "Statut", "Responsable", "Date limite"],
                        rows=[
                            ["Application Mobile", "{color:green}Terminé{/color}", "Sophie Martin", "15/03/2025"],
                            ["Refonte du site", "{color:blue}En cours{/color}", "Nicolas Durand", "30/05/2025"],
                            ["Intégration API", "{color:red}Retard{/color}", "Émilie Petit", "10/04/2025"],
                            ["Migration Cloud", "{color:orange}En attente{/color}", "Thomas Leroy", "22/06/2025"],
                            ["Audit Sécurité", "**Planifié**", "Claire Dubois", "05/07/2025"]
                        ]
                    )
                )
            )
        ]
    )


def create_comparison_table_slide():
    """Create a slide with a comparison table."""
    return Slide(
        id="comparison_table",
        title="Comparaison des Offres",
        layout_name="Titre et tableau",
        blocks=[
            SlideBlock(
                id="comparison_table_block",
                content=SlideContent(
                    content_type=ContentType.TABLE,
                    table=TableData(
                        headers=["Fonctionnalité", "Offre Basic", "Offre Premium", "Offre Enterprise", "style:grid"],
                        rows=[
                            ["Nombre d'utilisateurs", "5", "20", "Illimité"],
                            ["Stockage", "10 GB", "100 GB", "1 TB"],
                            ["Support", "E-mail", "E-mail + Téléphone", "24/7 Dédié"],
                            ["Sauvegardes", "Hebdomadaires", "Quotidiennes", "Temps réel"],
                            ["Personnalisation", "Limitée", "Standard", "Complète"],
                            ["Intégrations", "3 apps", "10 apps", "Illimitées"],
                            ["Prix mensuel", "29 €/mois", "99 €/mois", "299 €/mois"]
                        ]
                    )
                )
            )
        ]
    )


def create_mixed_content_slide():
    """Create a slide with both table and text content."""
    return Slide(
        id="mixed_content",
        title="Table avec Explication",
        layout_name="Titre et 3 colonnes",
        blocks=[
            SlideBlock(
                title="Résultats par trimestre",
                content=SlideContent(
                    content_type=ContentType.TABLE,
                    table=TableData(
                        headers=["Métrique", "T1", "T2", "T3", "T4"],
                        rows=[
                            ["Revenus", "1,2M€", "1,5M€", "1,8M€", "2,1M€"],
                            ["Coûts", "0,8M€", "0,9M€", "1,0M€", "1,1M€"],
                            ["Profit", "0,4M€", "0,6M€", "0,8M€", "1,0M€"]
                        ]
                    )
                )
            ),
            SlideBlock(
                id="text_block",
                title="Analyse",
                content=SlideContent(
                    content_type=ContentType.TEXT,
                    text="On observe une croissance constante des revenus et des profits tout au long de l'année, avec la plus forte augmentation au T4."
                )
            )
        ]
    )


def create_comparison_slide():
    """Create a slide with a comparison table."""
    return Slide(
        id="comparison",
        title="Comparaison des Offres",
        layout_name="Titre et tableau",
        blocks=[
            SlideBlock(
                id="comparison_block",
                content=SlideContent(
                    content_type=ContentType.TABLE,
                    table=TableData(
                        headers=["Fonctionnalité", "Offre Basic", "Offre Premium", "Offre Enterprise", "style:grid"],
                        rows=[
                            ["Utilisateurs", "5", "20", "Illimité"],
                            ["Stockage", "10 Go", "100 Go", "1 To"],
                            ["Support", "E-mail", "E-mail + Téléphone", "24/7 Dédié"],
                            ["Sauvegardes", "Hebdomadaires", "Quotidiennes", "Temps réel"],
                            ["API", "Non", "Limité", "Complet"],
                            ["Prix mensuel", "29€", "99€", "299€"]
                        ]
                    )
                )
            )
        ]
    )


def create_multi_column_slide():
    """Create a slide with multiple columns of content."""
    return Slide(
        id="multi_column",
        title="Points Clés par Département",
        layout_name="Titre et 3 colonnes",
        blocks=[
            SlideBlock(
                id="col1_block",
                title="Marketing",
                content=SlideContent(
                    content_type=ContentType.BULLET_POINTS,
                    bullet_points=[
                        "Augmentation des leads de 23%",
                        "Lancement réussi sur 3 nouveaux marchés",
                        "Budget optimisé de 15%",
                        "Nouveau site web : +45% de trafic"
                    ]
                )
            ),
            SlideBlock(
                id="col2_block",
                title="R&D",
                content=SlideContent(
                    content_type=ContentType.BULLET_POINTS,
                    bullet_points=[
                        "8 nouveaux brevets déposés",
                        "Prototype V2 prêt pour tests",
                        "Réduction du temps de dev de 30%",
                        "3 partenariats universitaires"
                    ]
                )
            ),
            SlideBlock(
                id="col3_block",
                title="Ventes",
                content=SlideContent(
                    content_type=ContentType.BULLET_POINTS,
                    bullet_points=[
                        "CA en hausse de 18% (YoY)",
                        "4 nouveaux grands comptes",
                        "Expansion internationale réussie",
                        "Taux de renouvellement : 92%"
                    ]
                )
            )
        ]
    )


def main():
    """Create and save a sample presentation with various table formats."""
    parser = argparse.ArgumentParser(description='Generate sample PowerPoint with tables')
    parser.add_argument('--output', '-o', default='example_tables.pptx', 
                      help='Output PowerPoint file path')
    parser.add_argument('--template', '-t', default='base_template.pptx',
                      help='Template PowerPoint file to use')
    
    args = parser.parse_args()
    
    # Check if template exists
    template_path = Path(args.template)
    if not template_path.exists():
        print(f"Template file {template_path} not found. Please specify a valid template file.")
        print("Using default template path: tests/fixtures/base_template.pptx")
        template_path = project_root / "tests" / "fixtures" / "base_template.pptx"
        if not template_path.exists():
            print("Default template not found either. Please ensure a template file exists.")
            return 1
    
    # Create slides
    simple_table_slide = create_simple_table_slide()
    styled_table_slide = create_styled_table_slide()
    formatted_table_slide = create_formatted_table_slide()
    comparison_slide = create_comparison_slide()
    mixed_content_slide = create_mixed_content_slide()
    multi_column_slide = create_multi_column_slide()
    
    # Create presentation
    presentation = Presentation(
        title="Exemples de Tables et Layouts",
        author="doc2pptx",
        description="Démonstration des fonctionnalités améliorées de tables et layouts",
        template_path=template_path,
        sections=[
            Section(
                title="Tables et Mise en Page",
                type=SectionType.CONTENT,
                slides=[
                    simple_table_slide,
                    styled_table_slide,
                    formatted_table_slide,
                    comparison_slide,
                    mixed_content_slide,
                    multi_column_slide
                ]
            )
        ]
    )
    
    # Create builder and build presentation
    builder = PPTBuilder(template_path=template_path)
    output_path = builder.build(presentation, args.output)
    
    print(f"Presentation created successfully: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())