#!/usr/bin/env python3
"""
Exemple d'utilisation du générateur d'images PowerPoint
"""

import os
from src.services.powerpoint_image_generator import PowerPointImageGenerator, Config

def main():
    """Exemple d'utilisation du générateur"""
    
    # Configuration de l'API DeepAI
    api_key = "your-deepai-api-key-here"  # Remplacez par votre vraie clé API
    
    # Chemin vers votre présentation PowerPoint
    input_file = "Messe du dimanche 15 janvier 2024.pptx"
    output_file = "Messe du dimanche 15 janvier 2024_with_images.pptx"
    
    try:
        print("🚀 Démarrage du générateur d'images PowerPoint")
        print(f"📁 Fichier d'entrée: {input_file}")
        print(f"📁 Fichier de sortie: {output_file}")
        
        # Initialiser le générateur
        generator = PowerPointImageGenerator(api_key=api_key)
        
        # Traiter la présentation
        success = generator.process_presentation(input_file, output_file)
        
        if success:
            print("✅ Présentation améliorée avec succès!")
            print(f"💾 Fichier sauvegardé: {output_file}")
        else:
            print("❌ Échec de l'amélioration de la présentation")
            
    except Exception as e:
        print(f"❌ Erreur: {str(e)}")

if __name__ == "__main__":
    main()