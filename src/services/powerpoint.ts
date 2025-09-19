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

    let slideNumber = 1;

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
      fontFace: 'Arial',
      bold: true
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

    // Numéro de slide
    titleSlide.addText(`${slideNumber}`, {
      x: 9.2,
      y: 7,
      w: 0.5,
      h: 0.3,
      fontSize: 12,
      color: liturgicalSeason.textColor,
      align: 'center',
      fontFace: 'Arial'
    });

    slideNumber++;

    // Slides dans l'ordre défini par l'utilisateur
    const sortedSlides = [...presentation.slideOrder].sort((a, b) => a.order - b.order);
    
    for (const slideItem of sortedSlides) {
      if (slideItem.type === 'reading') {
        const reading = presentation.readings.find(r => r.id === slideItem.id);
        if (reading) {
          const readingSlides = this.createReadingSlides(pptx, reading, liturgicalSeason, slideNumber);
          slideNumber += readingSlides;
        }
      } else if (slideItem.type === 'song') {
        const song = presentation.songs.find(s => s.id === slideItem.id);
        if (song) {
          const songSlides = this.createSongSlides(pptx, song, liturgicalSeason, slideNumber);
          slideNumber += songSlides;
        }
      }
    }

    // Télécharger le fichier
    await pptx.writeFile({ fileName: `${presentation.title}.pptx` });
  }

  private static createReadingSlides(
    pptx: PptxGenJS, 
    reading: LiturgyReading, 
    liturgicalSeason: any, 
    startSlideNumber: number
  ): number {
    let slideCount = 0;

    // Slide de titre pour la lecture
    const titleSlide = pptx.addSlide();
    titleSlide.background = { color: liturgicalSeason.backgroundColor };

    titleSlide.addText(reading.title, {
      x: 0.5,
      y: 2.5,
      w: 9,
      h: 1.2,
      fontSize: 36,
      color: liturgicalSeason.textColor,
      bold: true,
      align: 'center',
      fontFace: 'Arial'
    });

    titleSlide.addText(reading.reference, {
      x: 0.5,
      y: 4,
      w: 9,
      h: 0.8,
      fontSize: 24,
      color: liturgicalSeason.textColor,
      italic: true,
      align: 'center',
      fontFace: 'Arial'
    });

    // Numéro de slide
    titleSlide.addText(`${startSlideNumber + slideCount}`, {
      x: 9.2,
      y: 7,
      w: 0.5,
      h: 0.3,
      fontSize: 12,
      color: liturgicalSeason.textColor,
      align: 'center',
      fontFace: 'Arial'
    });

    slideCount++;

    // Diviser le texte en paragraphes
    const paragraphs = this.splitTextIntoParagraphs(reading.text);

    paragraphs.forEach((paragraph, index) => {
      const slide = pptx.addSlide();
      slide.background = { color: liturgicalSeason.backgroundColor };

      // Titre de la slide basé sur le contenu du paragraphe
      const slideTitle = this.generateSlideTitle(paragraph, reading.title, index + 1);
      
      slide.addText(slideTitle, {
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

      // Contenu formaté du paragraphe
      const formattedContent = this.formatParagraphContent(paragraph);
      
      slide.addText(formattedContent, {
        x: 0.8,
        y: 1.8,
        w: 8.4,
        h: 5,
        fontSize: 20,
        color: liturgicalSeason.textColor,
        align: 'left',
        fontFace: 'Arial',
        valign: 'top',
        lineSpacing: 28
      });

      // Numéro de slide
      slide.addText(`${startSlideNumber + slideCount}`, {
        x: 9.2,
        y: 7,
        w: 0.5,
        h: 0.3,
        fontSize: 12,
        color: liturgicalSeason.textColor,
        align: 'center',
        fontFace: 'Arial'
      });

      slideCount++;
    });

    return slideCount;
  }

  private static createSongSlides(
    pptx: PptxGenJS, 
    song: Song, 
    liturgicalSeason: any, 
    startSlideNumber: number
  ): number {
    let slideCount = 0;

    // Slide de titre pour le chant
    const titleSlide = pptx.addSlide();
    titleSlide.background = { color: liturgicalSeason.backgroundColor };

    titleSlide.addText(song.title, {
      x: 0.5,
      y: 2.5,
      w: 9,
      h: 1.2,
      fontSize: 32,
      color: liturgicalSeason.textColor,
      bold: true,
      align: 'center',
      fontFace: 'Arial'
    });

    if (song.author || song.melody) {
      const subtitle = [song.author, song.melody].filter(Boolean).join(' - ');
      titleSlide.addText(subtitle, {
        x: 0.5,
        y: 4,
        w: 9,
        h: 0.6,
        fontSize: 18,
        color: liturgicalSeason.textColor,
        italic: true,
        align: 'center',
        fontFace: 'Arial'
      });
    }

    // Numéro de slide
    titleSlide.addText(`${startSlideNumber + slideCount}`, {
      x: 9.2,
      y: 7,
      w: 0.5,
      h: 0.3,
      fontSize: 12,
      color: liturgicalSeason.textColor,
      align: 'center',
      fontFace: 'Arial'
    });

    slideCount++;

    // Diviser les paroles en couplets/strophes
    const verses = this.splitSongIntoVerses(song.lyrics);

    verses.forEach((verse, index) => {
      const slide = pptx.addSlide();
      slide.background = { color: liturgicalSeason.backgroundColor };

      // Titre de la slide
      const slideTitle = this.generateVerseTitle(verse, song.title, index + 1);
      
      slide.addText(slideTitle, {
        x: 0.5,
        y: 0.5,
        w: 9,
        h: 0.8,
        fontSize: 24,
        color: liturgicalSeason.textColor,
        bold: true,
        align: 'center',
        fontFace: 'Arial'
      });

      // Paroles formatées
      slide.addText(verse.trim(), {
        x: 1,
        y: 1.8,
        w: 8,
        h: 5,
        fontSize: 18,
        color: liturgicalSeason.textColor,
        align: 'center',
        fontFace: 'Arial',
        valign: 'top',
        lineSpacing: 24
      });

      // Numéro de slide
      slide.addText(`${startSlideNumber + slideCount}`, {
        x: 9.2,
        y: 7,
        w: 0.5,
        h: 0.3,
        fontSize: 12,
        color: liturgicalSeason.textColor,
        align: 'center',
        fontFace: 'Arial'
      });

      slideCount++;
    });

    return slideCount;
  }

  private static splitTextIntoParagraphs(text: string): string[] {
    // Diviser le texte en paragraphes basés sur les sauts de ligne doubles ou les points suivis d'espaces
    const paragraphs = text
      .split(/\n\s*\n|\.\s+(?=[A-Z])/g)
      .map(p => p.trim())
      .filter(p => p.length > 0);

    // Si un seul paragraphe très long, le diviser par phrases
    if (paragraphs.length === 1 && paragraphs[0].length > 400) {
      return this.splitLongParagraph(paragraphs[0]);
    }

    return paragraphs;
  }

  private static splitLongParagraph(paragraph: string): string[] {
    const sentences = paragraph.split(/(?<=[.!?])\s+/);
    const chunks: string[] = [];
    let currentChunk = '';

    for (const sentence of sentences) {
      if ((currentChunk + sentence).length > 300 && currentChunk.length > 0) {
        chunks.push(currentChunk.trim());
        currentChunk = sentence;
      } else {
        currentChunk += (currentChunk ? ' ' : '') + sentence;
      }
    }

    if (currentChunk.trim()) {
      chunks.push(currentChunk.trim());
    }

    return chunks;
  }

  private static splitSongIntoVerses(lyrics: string): string[] {
    // Diviser les paroles en couplets basés sur les lignes vides ou les marqueurs de refrain
    const verses = lyrics
      .split(/\n\s*\n|(?=R\/)|(?<=\n)(?=\d+\.)/g)
      .map(v => v.trim())
      .filter(v => v.length > 0);

    // Limiter chaque couplet à 6-8 lignes maximum
    const processedVerses: string[] = [];
    
    for (const verse of verses) {
      const lines = verse.split('\n');
      if (lines.length <= 8) {
        processedVerses.push(verse);
      } else {
        // Diviser les longs couplets
        for (let i = 0; i < lines.length; i += 6) {
          const chunk = lines.slice(i, i + 6).join('\n');
          processedVerses.push(chunk);
        }
      }
    }

    return processedVerses;
  }

  private static generateSlideTitle(paragraph: string, readingTitle: string, index: number): string {
    // Extraire les premiers mots significatifs pour créer un titre
    const words = paragraph.split(' ').slice(0, 6);
    let title = words.join(' ');
    
    // Nettoyer et raccourcir le titre
    title = title.replace(/[.!?].*$/, '');
    if (title.length > 50) {
      title = title.substring(0, 47) + '...';
    }

    return `${readingTitle} (${index})`;
  }

  private static generateVerseTitle(verse: string, songTitle: string, index: number): string {
    // Identifier le type de couplet
    if (verse.startsWith('R/') || verse.toLowerCase().includes('refrain')) {
      return `${songTitle} - Refrain`;
    } else if (/^\d+\./.test(verse)) {
      const match = verse.match(/^(\d+)\./);
      return `${songTitle} - Couplet ${match ? match[1] : index}`;
    } else {
      return `${songTitle} - Partie ${index}`;
    }
  }

  private static formatParagraphContent(paragraph: string): string {
    // Formater le contenu en points clés si le paragraphe est long
    if (paragraph.length > 200) {
      const sentences = paragraph.split(/(?<=[.!?])\s+/);
      if (sentences.length > 2) {
        return sentences.map(sentence => `• ${sentence.trim()}`).join('\n');
      }
    }
    
    return paragraph;
  }
}