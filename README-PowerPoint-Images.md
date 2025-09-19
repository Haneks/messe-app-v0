# 🎨 Générateur d'Images PowerPoint avec DeepAI

Ce script Python automatise la génération d'images contextuelles pour vos présentations PowerPoint liturgiques en utilisant l'API DeepAI.

## ✨ Fonctionnalités

- **🔍 Extraction automatique** des titres de chaque slide
- **🎨 Génération d'images** contextuelles avec l'API DeepAI
- **📐 Insertion intelligente** des images avec positionnement optimal
- **🛡️ Gestion d'erreurs** robuste avec retry automatique
- **📊 Logging détaillé** pour suivre le processus
- **🎯 Optimisation liturgique** avec prompts spécialisés

## 🚀 Installation

### 1. Installer les dépendances Python

```bash
pip install -r requirements-powerpoint.txt
```

### 2. Obtenir une clé API DeepAI

1. Créez un compte sur [DeepAI](https://deepai.org/)
2. Obtenez votre clé API depuis le dashboard
3. Configurez la variable d'environnement :

```bash
export DEEPAI_API_KEY="votre-cle-api-deepai"
```

## 📖 Utilisation

### Utilisation en ligne de commande

```bash
# Utilisation basique
python src/services/powerpoint-image-generator.py "presentation.pptx"

# Avec fichier de sortie personnalisé
python src/services/powerpoint-image-generator.py "presentation.pptx" -o "presentation_enhanced.pptx"

# Avec clé API spécifique
python src/services/powerpoint-image-generator.py "presentation.pptx" -k "votre-cle-api"

# Mode verbose
python src/services/powerpoint-image-generator.py "presentation.pptx" -v
```

### Utilisation programmatique

```python
from src.services.powerpoint_image_generator import PowerPointImageGenerator

# Initialiser le générateur
generator = PowerPointImageGenerator(api_key="votre-cle-api")

# Traiter une présentation
success = generator.process_presentation(
    input_path="presentation.pptx",
    output_path="presentation_with_images.pptx"
)

if success:
    print("✅ Images générées avec succès!")
```

## ⚙️ Configuration

### Variables d'environnement

```bash
# Clé API DeepAI (obligatoire)
export DEEPAI_API_KEY="votre-cle-api-deepai"

# Configuration optionnelle
export DEEPAI_TIMEOUT=30
export DEEPAI_MAX_RETRIES=3
```

### Configuration dans le code

```python
from src.services.powerpoint_image_generator import Config

# Modifier la configuration
Config.IMAGE_WIDTH = Inches(5)
Config.IMAGE_HEIGHT = Inches(4)
Config.IMAGE_POSITION_X = Inches(7)
Config.MAX_RETRIES = 5
```

## 🎯 Optimisations liturgiques

Le script reconnaît automatiquement le contenu liturgique et optimise les prompts :

### Mots-clés reconnus :
- **messe** → `catholic mass ceremony, church interior, altar`
- **évangile** → `gospel book, bible, religious scripture, holy light`
- **communion** → `holy communion, eucharist, chalice, bread and wine`
- **chant** → `church choir, religious music, hymn, spiritual singing`
- **noël** → `christmas, nativity, star, peaceful night`
- **pâques** → `easter, resurrection, sunrise, hope, new life`

### Exemple de transformation :
```
Titre: "Évangile selon Saint Matthieu"
Prompt généré: "Évangile selon Saint Matthieu, gospel book, bible, religious scripture, holy light, high quality, professional, clean, beautiful lighting, artistic composition"
```

## 📊 Fonctionnalités avancées

### Gestion des erreurs
- **Retry automatique** en cas d'échec API
- **Validation des images** téléchargées
- **Fallback gracieux** si une image échoue
- **Logging détaillé** pour le débogage

### Positionnement intelligent
- **Détection automatique** de la taille des slides
- **Évitement des zones de texte** existantes
- **Positionnement responsive** selon le contenu
- **Remplacement des images** existantes

### Performance
- **Téléchargement optimisé** des images
- **Cache temporaire** pour éviter les re-téléchargements
- **Nettoyage automatique** des fichiers temporaires
- **Respect des limites** de l'API

## 🔧 Dépannage

### Problèmes courants

#### Erreur "API key not configured"
```bash
export DEEPAI_API_KEY="votre-vraie-cle-api"
```

#### Erreur "Missing required libraries"
```bash
pip install python-pptx pillow requests
```

#### Timeout API
```python
Config.REQUEST_TIMEOUT = 60  # Augmenter le timeout
Config.MAX_RETRIES = 5       # Plus de tentatives
```

#### Images de mauvaise qualité
```python
# Modifier le prompt de base dans generate_image_prompt()
final_prompt = f"{enhanced_prompt}, high resolution, professional photography, detailed, artistic"
```

### Logs de débogage

Le script génère des logs détaillés dans `powerpoint_image_generator.log` :

```
2024-01-15 10:30:15 - INFO - 📄 Slide 1: 'Messe du dimanche'
2024-01-15 10:30:16 - INFO - 🎨 Generated prompt for slide 1: 'Messe du dimanche, catholic mass ceremony...'
2024-01-15 10:30:20 - INFO - ✅ Downloaded image for slide 1: /tmp/slide_1_image.jpg
2024-01-15 10:30:21 - INFO - ✅ Inserted image into slide 1
```

## 🎨 Personnalisation des prompts

Vous pouvez personnaliser les prompts en modifiant la méthode `generate_image_prompt()` :

```python
def generate_image_prompt(self, slide_title: str, slide_index: int) -> str:
    # Vos règles personnalisées ici
    if "baptême" in slide_title.lower():
        return f"{slide_title}, baptism ceremony, water, dove, holy spirit, peaceful"
    
    # Logique par défaut...
```

## 📈 Intégration avec l'application

Le script s'intègre parfaitement avec votre application liturgique existante :

```typescript
import { PowerPointImageService } from './services/powerpoint-integration';

// Dans votre composant PreviewPanel
const handleExportWithImages = async () => {
  const config = {
    apiKey: process.env.DEEPAI_API_KEY,
    ...PowerPointImageService.getDefaultConfig()
  };
  
  const success = await PowerPointImageService.enhanceWithImages(presentation, config);
  
  if (success) {
    alert('✅ Présentation avec images exportée!');
  }
};
```

## 💡 Conseils d'utilisation

1. **Testez d'abord** avec une petite présentation
2. **Vérifiez votre quota** API DeepAI
3. **Sauvegardez** vos présentations originales
4. **Ajustez les positions** selon vos besoins
5. **Personnalisez les prompts** pour votre contexte

## 🆘 Support

Pour toute question ou problème :
1. Consultez les logs détaillés
2. Vérifiez votre clé API DeepAI
3. Testez avec le mode verbose (`-v`)
4. Vérifiez les permissions de fichiers

---

**🎉 Votre présentation liturgique sera maintenant enrichie d'images contextuelles générées automatiquement !**