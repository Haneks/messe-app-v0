import { LiturgyReading, AELFResponse } from '../types/liturgy';

// Service pour récupérer les textes liturgiques de l'AELF
export class AELFService {
  private static readonly BASE_URL = 'https://api.aelf.org';

  static async getReadingsForDate(date: string): Promise<AELFResponse> {
    try {
      // Format de date pour l'API AELF (YYYY/MM/DD)
      const formattedDate = date.replace(/-/g, '/');
      const url = `${this.BASE_URL}/v1/messes/${formattedDate}`;
      
      console.log('Appel API AELF:', url);
      
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        console.warn(`Erreur API AELF (${response.status}):`, response.statusText);
        return this.getMockReadings(date);
      }

      const data = await response.json();
      console.log('Réponse API AELF:', data);
      
      return this.parseAELFResponse(data, date);
    } catch (error) {
      console.error('Erreur lors de la récupération des textes AELF:', error);
      // Fallback vers les données simulées en cas d'erreur
      return this.getMockReadings(date);
    }
  }

  private static parseAELFResponse(data: any, date: string): AELFResponse {
    const readings: AELFResponse['readings'] = {};

    try {
      // Première lecture
      if (data.lecture_1) {
        readings.first_reading = {
          id: 'first-' + date,
          title: 'Première lecture',
          reference: data.lecture_1.reference || '',
          text: this.cleanText(data.lecture_1.text || data.lecture_1.contenu || ''),
          type: 'first_reading'
        };
      }

      // Psaume
      if (data.psaume) {
        readings.psalm = {
          id: 'psalm-' + date,
          title: 'Psaume responsorial',
          reference: data.psaume.reference || '',
          text: this.cleanText(data.psaume.text || data.psaume.contenu || ''),
          type: 'psalm'
        };
      }

      // Deuxième lecture (si présente)
      if (data.lecture_2) {
        readings.second_reading = {
          id: 'second-' + date,
          title: 'Deuxième lecture',
          reference: data.lecture_2.reference || '',
          text: this.cleanText(data.lecture_2.text || data.lecture_2.contenu || ''),
          type: 'second_reading'
        };
      }

      // Évangile
      if (data.evangile) {
        readings.gospel = {
          id: 'gospel-' + date,
          title: 'Évangile',
          reference: data.evangile.reference || '',
          text: this.cleanText(data.evangile.text || data.evangile.contenu || ''),
          type: 'gospel'
        };
      }

      // Si aucune lecture n'a été trouvée, utiliser les données simulées
      if (Object.keys(readings).length === 0) {
        console.warn('Aucune lecture trouvée dans la réponse AELF, utilisation des données simulées');
        return this.getMockReadings(date);
      }

    } catch (parseError) {
      console.error('Erreur lors du parsing de la réponse AELF:', parseError);
      return this.getMockReadings(date);
    }

    return { readings };
  }

  private static cleanText(text: string): string {
    if (!text) return '';
    
    // Nettoyer le texte des balises HTML et caractères indésirables
    return text
      .replace(/<[^>]*>/g, '') // Supprimer les balises HTML
      .replace(/&nbsp;/g, ' ') // Remplacer les espaces insécables
      .replace(/&amp;/g, '&') // Remplacer les entités HTML
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/\s+/g, ' ') // Normaliser les espaces multiples
      .trim();
  }

  private static getMockReadings(date: string): AELFResponse {
    return {
      readings: {
        first_reading: {
          id: 'first-' + date,
          title: 'Première lecture',
          reference: 'Is 55, 10-11',
          text: 'Ainsi parle le Seigneur : « La pluie et la neige qui descendent des cieux n\'y retournent pas sans avoir abreuvé la terre, sans l\'avoir fécondée et l\'avoir fait germer, donnant la semence au semeur et le pain à celui qui doit manger ; ainsi ma parole, qui sort de ma bouche, ne me reviendra pas sans résultat, sans avoir fait ce qui me plaît, sans avoir accompli sa mission. »',
          type: 'first_reading'
        },
        psalm: {
          id: 'psalm-' + date,
          title: 'Psaume responsorial',
          reference: 'Ps 64',
          text: 'Tu visites la terre et tu l\'abreuves, tu la combles de richesses ; les ruisseaux de Dieu regorgent d\'eau : tu prépares les moissons. Ainsi tu prépares la terre, tu arroses les sillons ; tu aplanis le sol, tu le détrempes sous les pluies, tu bénis les semailles.',
          type: 'psalm'
        },
        gospel: {
          id: 'gospel-' + date,
          title: 'Évangile',
          reference: 'Mt 13, 1-23',
          text: 'Ce jour-là, Jésus était sorti de la maison, et il était assis au bord de la mer. Auprès de lui se rassemblèrent des foules si grandes qu\'il monta dans une barque où il s\'assit ; toute la foule se tenait sur le rivage. Il leur dit beaucoup de choses en paraboles : « Voici que le semeur sortit pour semer. Comme il semait, des grains sont tombés au bord du chemin, et les oiseaux sont venus tout manger. »',
          type: 'gospel'
        }
      }
    };
  }
}