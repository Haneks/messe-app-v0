import PptxGenJS from 'pptxgenjs';
import { LiturgyPresentation, LiturgyReading, Song, SlideItem } from '../types/liturgy';
import { getLiturgicalSeason } from '../utils/liturgicalColors';

export class PowerPointService {
  static async exportPresentation(presentation: LiturgyPresentation): Promise<void> {
    const pptx = new PptxGenJS();
    
    // Déterminer les couleurs liturgiques selon la date
    const liturgicalSeason = getLiturgicalSeason(new Date(presentation.date));
    
    // Configuration de base
    pptx.author = 'Application Liturgique';
    pptx.title = presentation.title;
    pptx.subject = 'Présentation pour la messe';

    // Slide de titre
    const titleSlide = pptx.addSlide();
    titleSlide.background = { color: '1E40AF' };
    
    titleSlide.addText(presentation.title, {
      x: 0.5,
      y: 2,
      w: 9,
      h: 1.5,
      fontSize: 44,
      color: liturgicalSeason.textColor,
      align: 'center',
      fontFace: 'Arial'
    });

    titleSlide.addText(new Date(presentation.date).toLocaleDateString('fr-FR', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    }), {
      x: 0.5,
      y: 4,
      w: 9,
      h: 1,
      fontSize: 24,
      color: liturgicalSeason.textColor,
      align: 'center',
      fontFace: 'Arial'
    });

    // Slides dans l'ordre défini par l'utilisateur
    const sortedSlides = [...presentation.slideOrder].sort((a, b) => a.order - b.order);
    
    for (const slideItem of sortedSlides) {
      if (slideItem.type === 'reading') {
        const reading = presentation.readings.find(r => r.id === slideItem.id);
        if (reading) {
          const slide = pptx.addSlide();
          slide.background = { color: liturgicalSeason.backgroundColor };

          // Titre de la lecture
          slide.addText(reading.title, {
            x: 0.5,
            y: 0.5,
            w: 9,
            h: 0.8,
            fontSize: 32,
            color: liturgicalSeason.textColor,
            bold: true,
            align: 'center',
            fontFace: 'Arial'
          });

          // Référence
          slide.addText(reading.reference, {
            x: 0.5,
            y: 1.4,
            w: 9,
            h: 0.6,
            fontSize: 20,
            color: liturgicalSeason.textColor,
            italic: true,
            align: 'center',
            fontFace: 'Arial'
          });

          // Texte
          slide.addText(reading.text, {
            x: 0.8,
            y: 2.2,
            w: 8.4,
            h: 4.8,
            fontSize: 18,
            color: liturgicalSeason.textColor,
            align: 'left',
            fontFace: 'Arial',
            valign: 'top'
          });
        }
      } else if (slideItem.type === 'song') {
        const song = presentation.songs.find(s => s.id === slideItem.id);
        if (song) {
          const slide = pptx.addSlide();
          slide.background = { color: liturgicalSeason.backgroundColor };

          // Titre du chant
          slide.addText(song.title, {
            x: 0.5,
            y: 0.5,
            w: 9,
            h: 0.8,
            fontSize: 28,
            color: liturgicalSeason.textColor,
            bold: true,
            align: 'center',
            fontFace: 'Arial'
          });

          // Paroles
          slide.addText(song.lyrics, {
            x: 0.8,
            y: 1.5,
            w: 8.4,
            h: 5.5,
            fontSize: 16,
            color: liturgicalSeason.textColor,
            align: 'left',
            fontFace: 'Arial',
            valign: 'top'
          });
        }
      }
    }

    // Télécharger le fichier
    await pptx.writeFile({ fileName: `${presentation.title}.pptx` });
  }
}