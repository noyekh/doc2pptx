"""
Script pour corriger le problème de types de section dans core/models.py.
"""

import re
from pathlib import Path

def fix_section_types(models_file_path):
    # Lire le contenu du fichier
    with open(models_file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Identifier la classe SectionType
    section_type_pattern = r'class SectionType\(str, Enum\):(.*?)(?=class|$)'
    section_type_match = re.search(section_type_pattern, content, re.DOTALL)
    
    if not section_type_match:
        print("Classe SectionType non trouvée dans le fichier.")
        return False
    
    # Extraire la définition actuelle
    section_type_def = section_type_match.group(1)
    
    # Vérifier si les types supplémentaires sont déjà présents
    if "SECTION_HEADER" in section_type_def:
        print("Les types supplémentaires semblent déjà être présents.")
        return False
    
    # Ajouter les types supplémentaires
    additional_types = """
    # Types supplémentaires pour les tests
    SECTION_HEADER = "section_header"
    BULLET_LIST = "bullet_list"
    CHART = "chart"
    TEXT_BLOCKS = "text_blocks"
    IMAGE_RIGHT = "image_right"
    TWO_COLUMN = "two_column"
    TABLE = "table"
    IMAGE_LEFT = "image_left"
    HEAT_MAP = "heat_map"
    QUOTE = "quote"
    NUMBERED_LIST = "numbered_list"
    THANK_YOU = "thank_you"
    """
    
    # Insérer les types supplémentaires juste avant la fermeture de la classe
    last_existing_type_pos = section_type_def.rfind('=')
    if last_existing_type_pos == -1:
        print("Format inattendu pour la classe SectionType.")
        return False
    
    # Trouver la position après le dernier type existant
    last_line_end = section_type_def.find('\n', last_existing_type_pos)
    if last_line_end == -1:
        last_line_end = len(section_type_def)
    
    # Créer la nouvelle définition
    new_section_type_def = section_type_def[:last_line_end] + additional_types + section_type_def[last_line_end:]
    
    # Remplacer l'ancienne définition par la nouvelle
    new_content = content.replace(section_type_match.group(1), new_section_type_def)
    
    # Écrire le contenu modifié dans le fichier
    with open(models_file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print(f"Le fichier {models_file_path} a été mis à jour avec succès.")
    return True

if __name__ == "__main__":
    # Déterminer l'emplacement du fichier models.py
    current_dir = Path.cwd()
    models_file = current_dir / "core" / "models.py"
    
    if not models_file.exists():
        # Essayer de trouver le fichier dans le répertoire parent
        models_file = current_dir.parent / "core" / "models.py"
    
    if not models_file.exists():
        # Essayer avec le chemin doc2pptx/core/models.py
        models_file = current_dir / "doc2pptx" / "core" / "models.py"
    
    if not models_file.exists():
        # Essayer avec le chemin src/doc2pptx/core/models.py
        models_file = current_dir / "src" / "doc2pptx" / "core" / "models.py"
    
    if not models_file.exists():
        print("Impossible de trouver le fichier models.py. Veuillez spécifier le chemin en argument.")
        import sys
        if len(sys.argv) > 1:
            models_file = Path(sys.argv[1])
            if not models_file.exists():
                print(f"Le fichier {models_file} n'existe pas.")
                sys.exit(1)
        else:
            sys.exit(1)
    
    fix_section_types(models_file)