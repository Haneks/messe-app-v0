import React from 'react';
import { Eye, Download, Image, Loader2 } from 'lucide-react';
import { LiturgyPresentation, SlideItem } from '../types/liturgy';
import { PowerPointService } from '../services/powerpoint';
import { PowerPointImageService } from '../services/powerpoint-integration';

interface PreviewPanelProps {
  presentation: LiturgyPresentation | null;
}

export default function PreviewPanel({ presentation }: PreviewPanelProps) {
  const [exporting, setExporting] = React.useState(false);
  const [exportingWithImages, setExportingWithImages] = React.useState(false);
  const [exportProgress, setExportProgress] = React.useState({ step: '', progress: 0 });
  const [showImageOptions, setShowImageOptions] = React.useState(false);
  const [apiKey, setApiKey] = React.useState('');

  const handleExport = async () => {
    if (!presentation) return;
    
    setExporting(true);
    try {
      await PowerPointService.exportPresentation(presentation);
    } catch (error) {
      console.error('Erreur lors de l\'export:', error);
      alert('Erreur lors de l\'export PowerPoint');
    } finally {
      setExporting(false);
    }
  };

  const handleExportWithImages = async () => {
    if (!presentation) return;
    
    // Vérifier la clé API
    if (!apiKey || !PowerPointImageService.validateApiKey(apiKey)) {
      alert('Veuillez entrer une clé API DeepAI valide');
      setShowImageOptions(true);
      return;
    }
    
    setExportingWithImages(true);
    setExportProgress({ step: 'Initialisation...', progress: 10 });
    
    try {
      // Étape 1: Export PowerPoint de base
      setExportProgress({ step: 'Création de la présentation de base...', progress: 20 });
      await PowerPointService.exportPresentation(presentation);
      
      // Étape 2: Configuration pour la génération d'images
      setExportProgress({ step: 'Configuration de la génération d\'images...', progress: 30 });
      const config = {
        apiKey: apiKey,
        ...PowerPointImageService.getDefaultConfig()
      };
      
      // Étape 3: Génération et insertion des images
      setExportProgress({ step: 'Génération des images contextuelles...', progress: 50 });
      const success = await PowerPointImageService.enhanceWithImages(presentation, config);
      
      if (success) {
        setExportProgress({ step: 'Finalisation de l\'export...', progress: 90 });
        
        // Petit délai pour montrer la progression
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        setExportProgress({ step: 'Export terminé !', progress: 100 });
        
        // Afficher le succès pendant 2 secondes
        setTimeout(() => {
          setExportProgress({ step: '', progress: 0 });
        }, 2000);
        
      } else {
        throw new Error('Échec de la génération d\'images');
      }
      
    } catch (error) {
      console.error('Erreur lors de l\'export avec images:', error);
      alert('Erreur lors de la génération d\'images. Export PowerPoint standard effectué.');
    } finally {
      setExportingWithImages(false);
    }
  };

  const handleApiKeySubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (PowerPointImageService.validateApiKey(apiKey)) {
      setShowImageOptions(false);
      handleExportWithImages();
    } else {
      alert('Clé API invalide. Veuillez vérifier votre clé DeepAI.');
    }
  };

  if (!presentation) {
    return (
      <div className="bg-white rounded-lg shadow-md p-6">
        <div className="flex items-center gap-3 mb-4">
          <Eye className="w-6 h-6 text-blue-700" />
          <h2 className="text-xl font-semibold text-gray-800">Aperçu</h2>
        </div>
        
        <div className="text-center py-12 text-gray-500">
          <Eye className="w-16 h-16 mx-auto mb-4 opacity-30" />
          <p>L'aperçu de votre présentation apparaîtra ici</p>
        </div>
      </div>
    );
  }

  // Calculer le nombre total de slides (titre + slides ordonnées)
  const totalSlides = 1 + presentation.slideOrder.length;

  const getSlideInfo = (slideItem: SlideItem, index: number) => {
    if (slideItem.type === 'reading') {
      const reading = presentation.readings.find(r => r.id === slideItem.id);
      return reading ? {
        title: reading.title,
        color: 'bg-green-600'
      } : null;
    } else {
      const song = presentation.songs.find(s => s.id === slideItem.id);
      return song ? {
        title: song.title,
        color: 'bg-amber-600'
      } : null;
    }
  };

  // Trier les slides par ordre
  const sortedSlides = [...presentation.slideOrder].sort((a, b) => a.order - b.order);

  return (
    <div className="bg-white rounded-lg shadow-md p-6">
      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center gap-3">
          <Eye className="w-6 h-6 text-blue-700" />
          <h2 className="text-xl font-semibold text-gray-800">Aperçu</h2>
        </div>
        
        <div className="flex gap-2">
          <button
            onClick={handleExport}
            disabled={exporting || exportingWithImages}
            className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 disabled:opacity-50 transition-colors"
          >
            <Download className="w-4 h-4" />
            {exporting ? 'Export...' : 'Export Simple'}
          </button>
          
          <button
            onClick={() => setShowImageOptions(true)}
            disabled={exporting || exportingWithImages}
            className="flex items-center gap-2 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 disabled:opacity-50 transition-colors"
          >
            {exportingWithImages ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <Image className="w-4 h-4" />
            )}
            {exportingWithImages ? 'Génération...' : 'Export + Images'}
          </button>
        </div>
      </div>

      {/* Barre de progression pour l'export avec images */}
      {exportingWithImages && (
        <div className="mb-6 bg-purple-50 border border-purple-200 rounded-lg p-4">
          <div className="flex items-center gap-3 mb-3">
            <Loader2 className="w-5 h-5 text-purple-600 animate-spin" />
            <span className="font-medium text-purple-800">Génération d'images en cours...</span>
          </div>
          
          <div className="space-y-2">
            <div className="flex justify-between text-sm text-purple-700">
              <span>{exportProgress.step}</span>
              <span>{exportProgress.progress}%</span>
            </div>
            <div className="w-full bg-purple-200 rounded-full h-2">
              <div 
                className="bg-purple-600 h-2 rounded-full transition-all duration-500 ease-out"
                style={{ width: `${exportProgress.progress}%` }}
              ></div>
            </div>
          </div>
          
          <p className="text-sm text-purple-600 mt-2">
            ⏳ Génération d'images contextuelles avec l'IA... Cela peut prendre quelques minutes.
          </p>
        </div>
      )}

      {/* Modal de configuration API */}
      {showImageOptions && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 max-w-md w-full mx-4">
            <h3 className="text-lg font-semibold text-gray-800 mb-4">
              Configuration de génération d'images
            </h3>
            
            <form onSubmit={handleApiKeySubmit} className="space-y-4">
              <div>
                <label htmlFor="deepai-key" className="block text-sm font-medium text-gray-700 mb-2">
                  Clé API DeepAI *
                </label>
                <input
                  type="password"
                  id="deepai-key"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder="Entrez votre clé API DeepAI"
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-purple-500 focus:border-transparent"
                  required
                />
                <p className="text-xs text-gray-500 mt-1">
                  Obtenez votre clé API sur <a href="https://deepai.org/" target="_blank" rel="noopener noreferrer" className="text-purple-600 hover:underline">deepai.org</a>
                </p>
              </div>
              
              <div className="bg-blue-50 p-3 rounded-lg">
                <h4 className="font-medium text-blue-800 mb-2">🎨 Fonctionnalités :</h4>
                <ul className="text-sm text-blue-700 space-y-1">
                  <li>• Images contextuelles pour chaque slide</li>
                  <li>• Optimisation pour le contenu liturgique</li>
                  <li>• Positionnement automatique intelligent</li>
                  <li>• Qualité professionnelle</li>
                </ul>
              </div>
              
              <div className="flex gap-3">
                <button
                  type="submit"
                  className="flex-1 px-4 py-2 bg-purple-600 text-white rounded-md hover:bg-purple-700 transition-colors"
                >
                  Générer les images
                </button>
                <button
                  type="button"
                  onClick={() => setShowImageOptions(false)}
                  className="px-4 py-2 bg-gray-300 text-gray-700 rounded-md hover:bg-gray-400 transition-colors"
                >
                  Annuler
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      <div className="space-y-4">
        <div className="bg-blue-50 p-4 rounded-lg border border-blue-200">
          <h3 className="font-semibold text-blue-800 mb-2">{presentation.title}</h3>
          <p className="text-blue-700 text-sm">
            {new Date(presentation.date).toLocaleDateString('fr-FR', {
              weekday: 'long',
              year: 'numeric',
              month: 'long',
              day: 'numeric'
            })}
          </p>
          <p className="text-blue-600 text-sm mt-2">
            {totalSlides} diapositive{totalSlides > 1 ? 's' : ''}
          </p>
        </div>

        <div className="space-y-3">
          <h4 className="font-medium text-gray-800">Contenu de la présentation :</h4>
          
          <div className="grid grid-cols-1 gap-2">
            <div className="flex items-center gap-3 p-3 bg-gray-50 rounded-lg">
              <div className="w-8 h-6 bg-blue-600 rounded flex items-center justify-center">
                <span className="text-white text-xs font-bold">1</span>
              </div>
              <span className="text-gray-700">Page de titre</span>
            </div>

            {sortedSlides.map((slideItem, index) => {
              const slideInfo = getSlideInfo(slideItem, index);
              if (!slideInfo) return null;

              return (
                <div key={slideItem.id} className="flex items-center gap-3 p-3 bg-gray-50 rounded-lg">
                  <div className={`w-8 h-6 ${slideInfo.color} rounded flex items-center justify-center`}>
                    <span className="text-white text-xs font-bold">{index + 2}</span>
                  </div>
                  <span className="text-gray-700">{slideInfo.title}</span>
                </div>
              );
            })}
          </div>
        </div>

        {sortedSlides.length > 0 && (
          <div className="bg-green-50 p-4 rounded-lg border border-green-200">
            <p className="text-green-800 text-sm">
              ✓ Présentation prête à être exportée en PowerPoint
            </p>
          </div>
        )}
      </div>
    </div>
  );
}