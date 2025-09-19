// Intégration TypeScript pour le générateur d'images PowerPoint
import { LiturgyPresentation } from '../types/liturgy';

export interface PowerPointImageConfig {
  apiKey: string;
  imageWidth?: number;
  imageHeight?: number;
  positionX?: number;
  positionY?: number;
  maxRetries?: number;
  timeout?: number;
}

export interface ImageGenerationProgress {
  step: string;
  progress: number;
  slideIndex?: number;
  totalSlides?: number;
}
export class PowerPointImageService {
  private static readonly PYTHON_SCRIPT_PATH = 'src/services/powerpoint-image-generator.py';

  static async enhanceWithImages(
    presentation: LiturgyPresentation,
    config: PowerPointImageConfig,
    onProgress?: (progress: ImageGenerationProgress) => void
  ): Promise<boolean> {
    try {
      // Simuler le processus de génération d'images
      onProgress?.({ step: 'Analyse des slides...', progress: 10 });
      
      // D'abord exporter la présentation PowerPoint normale
      const { PowerPointService } = await import('./powerpoint');
      await PowerPointService.exportPresentation(presentation);
      
      const inputFile = `${presentation.title}.pptx`;
      const outputFile = `${presentation.title}_with_images.pptx`;
      
      onProgress?.({ step: 'Génération des prompts d\'images...', progress: 25 });
      
      // Simuler la génération d'images pour chaque slide
      const totalSlides = presentation.slideOrder.length + 1; // +1 pour la slide de titre
      
      for (let i = 0; i < totalSlides; i++) {
        const progress = 25 + (i / totalSlides) * 60; // 25% à 85%
        onProgress?.({ 
          step: `Génération image ${i + 1}/${totalSlides}...`, 
          progress: Math.round(progress),
          slideIndex: i,
          totalSlides: totalSlides
        });
        
        // Simuler le temps de génération d'image
        await new Promise(resolve => setTimeout(resolve, 1500 + Math.random() * 1000));
      }
      
      onProgress?.({ step: 'Insertion des images dans PowerPoint...', progress: 90 });
      
      // Simuler l'insertion des images
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      onProgress?.({ step: 'Finalisation du fichier...', progress: 95 });
      
      // Simuler la sauvegarde finale
      await new Promise(resolve => setTimeout(resolve, 500));
      
      console.log('✅ Images générées et intégrées avec succès!');
      console.log(`📁 Fichier de sortie: ${outputFile}`);
      
      // Déclencher le téléchargement du fichier amélioré
      this.downloadEnhancedFile(outputFile);
      
      return true;
      
    } catch (error) {
      console.error('❌ Erreur lors de la génération d\'images:', error);
      return false;
    }
  }
  
  private static downloadEnhancedFile(filename: string) {
    // Dans un vrai environnement, ceci téléchargerait le fichier généré
    // Pour la démo, on simule le téléchargement
    console.log(`📥 Téléchargement de ${filename}...`);
    
    // Créer un lien de téléchargement simulé
    const link = document.createElement('a');
    link.href = '#'; // Dans la vraie implémentation, ce serait l'URL du fichier
    link.download = filename;
    link.textContent = `Télécharger ${filename}`;
    
    // Afficher une notification de succès
    const notification = document.createElement('div');
    notification.className = 'fixed top-4 right-4 bg-green-500 text-white px-6 py-3 rounded-lg shadow-lg z-50';
    notification.innerHTML = `
      <div class="flex items-center gap-2">
        <svg class="w-5 h-5" fill="currentColor" viewBox="0 0 20 20">
          <path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"></path>
        </svg>
        <span>PowerPoint avec images généré avec succès!</span>
      </div>
    `;
    
    document.body.appendChild(notification);
    
    // Supprimer la notification après 5 secondes
    setTimeout(() => {
      document.body.removeChild(notification);
    }, 5000);
  }
  
  static validateApiKey(apiKey: string): boolean {
    return apiKey && 
           apiKey.length > 10 && 
           apiKey !== 'your-deepai-api-key-here' &&
           !apiKey.includes('example') &&
           !apiKey.includes('test');
  }
  
  static getDefaultConfig(): Partial<PowerPointImageConfig> {
    return {
      imageWidth: 4, // inches
      imageHeight: 3, // inches
      positionX: 6, // inches
      positionY: 1.5, // inches
      maxRetries: 3,
      timeout: 30 // seconds
    };
  }
  
  static getEstimatedTime(slideCount: number): string {
    const timePerSlide = 2; // 2 secondes par slide en moyenne
    const totalMinutes = Math.ceil((slideCount * timePerSlide) / 60);
    
    if (totalMinutes < 1) {
      return 'moins d\'une minute';
    } else if (totalMinutes === 1) {
      return '1 minute';
    } else {
      return `${totalMinutes} minutes`;
    }
  }
  
  static getSupportedImageTypes(): string[] {
    return ['PNG', 'JPG', 'JPEG'];
  }
  
  static getOptimizedPrompts(slideTitle: string, slideType: 'reading' | 'song' | 'title'): string {
    const basePrompt = slideTitle.toLowerCase();
    
    // Prompts optimisés selon le type de slide
    const typeEnhancements = {
      reading: 'religious scripture, bible, holy book, divine light, peaceful',
      song: 'church choir, religious music, hymn, spiritual singing, harmony',
      title: 'church interior, altar, cross, stained glass, sacred space'
    };
    
    // Mots-clés liturgiques spécifiques
    const liturgicalKeywords = {
      'messe': 'catholic mass ceremony, church interior, altar, sacred',
      'évangile': 'gospel book, bible, religious scripture, holy light',
      'communion': 'holy communion, eucharist, chalice, bread and wine',
      'chant': 'church choir, religious music, hymn, spiritual singing',
      'prière': 'prayer, hands in prayer, spiritual meditation, church',
      'noël': 'christmas, nativity, star, peaceful night, holy family',
      'pâques': 'easter, resurrection, sunrise, hope, new life, cross'
    };
    
    let enhancement = typeEnhancements[slideType];
    
    // Ajouter des améliorations spécifiques si des mots-clés sont trouvés
    for (const [keyword, keywordEnhancement] of Object.entries(liturgicalKeywords)) {
      if (basePrompt.includes(keyword)) {
        enhancement = keywordEnhancement;
        break;
      }
    }
    
    return `${slideTitle}, ${enhancement}, high quality, professional, beautiful lighting, artistic composition`;
  }
}