import React, { useState } from 'react';
import { Music, Plus, Trash2, Edit, ChevronUp, ChevronDown, BookOpen } from 'lucide-react';
import { Song } from '../types/liturgy';
import { songLibrary } from '../data/songLibrary';

interface SongsPanelProps {
  songs: Song[];
  onAddSong: (song: Omit<Song, 'id'>) => void;
  onUpdateSong: (id: string, song: Omit<Song, 'id'>) => void;
  onDeleteSong: (id: string) => void;
  onReorderSong: (id: string, direction: 'up' | 'down') => void;
}

export default function SongsPanel({ songs, onAddSong, onUpdateSong, onDeleteSong, onReorderSong }: SongsPanelProps) {
  const [showAddForm, setShowAddForm] = useState(false);
  const [showLibrary, setShowLibrary] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [formData, setFormData] = useState({
    title: '',
    lyrics: '',
    author: '',
    melody: '',
    category: 'other',
    order: 0
  });

  const categories = [
    { value: 'entrance', label: 'Entrée' },
    { value: 'kyrie', label: 'Kyrie' },
    { value: 'gloria', label: 'Gloria' },
    { value: 'offertory', label: 'Offertoire' },
    { value: 'sanctus', label: 'Sanctus' },
    { value: 'communion', label: 'Communion' },
    { value: 'final', label: 'Sortie' },
    { value: 'other', label: 'Autre' }
  ];

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    if (editingId) {
      onUpdateSong(editingId, formData);
      setEditingId(null);
    } else {
      onAddSong(formData);
    }
    
    // Cette fonction est maintenant gérée par le composant SlideOrderPanel
    // On la garde pour la compatibilité mais elle ne fait plus rien
    console.log('Réorganisation des chants maintenant gérée par SlideOrderPanel');
    
    setFormData({
      title: '',
      lyrics: '',
      author: '',
      melody: '',
      category: 'other',
      order: 0
    });
    setShowAddForm(false);
  };

  const handleAddFromLibrary = (songTemplate: typeof songLibrary[0]) => {
    onAddSong({
      title: songTemplate.title,
      lyrics: songTemplate.lyrics,
      author: songTemplate.author,
      melody: songTemplate.melody,
      category: songTemplate.category,
      order: 0
    });
    setShowLibrary(false);
  };

  const handleEdit = (song: Song) => {
    setFormData({
      title: song.title,
      lyrics: song.lyrics,
      author: song.author || '',
      melody: song.melody || '',
      category: song.category,
      order: song.order
    });
    setEditingId(song.id);
    setShowAddForm(true);
  };

  const handleCancel = () => {
    setShowAddForm(false);
    setEditingId(null);
    setFormData({
      title: '',
      lyrics: '',
      author: '',
      melody: '',
      category: 'other',
      order: 0
    });
  };

  const getCategoryLabel = (category: string) => {
    return categories.find(cat => cat.value === category)?.label || category;
  };

  // Trier les chants par ordre
  const sortedSongs = [...songs].sort((a, b) => a.order - b.order);
  return (
    <div className="bg-white rounded-lg shadow-md p-6 mb-6">
      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center gap-3">
          <Music className="w-6 h-6 text-blue-700" />
          <h2 className="text-xl font-semibold text-gray-800">Chants liturgiques</h2>
        </div>
        
        <div className="flex gap-2">
          <button
            onClick={() => setShowLibrary(true)}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
          >
            <BookOpen className="w-4 h-4" />
            Bibliothèque
          </button>
          <button
            onClick={() => setShowAddForm(true)}
            className="flex items-center gap-2 px-4 py-2 bg-amber-600 text-white rounded-lg hover:bg-amber-700 transition-colors"
          >
            <Plus className="w-4 h-4" />
            Nouveau chant
          </button>
        </div>
      </div>

      {showLibrary && (
        <div className="mb-6 bg-blue-50 rounded-lg p-4 border border-blue-200">
          <div className="flex items-center justify-between mb-4">
            <h3 className="font-medium text-blue-800">Bibliothèque de chants liturgiques</h3>
            <button
              onClick={() => setShowLibrary(false)}
              className="text-blue-600 hover:text-blue-800 transition-colors"
            >
              ✕
            </button>
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-h-96 overflow-y-auto">
            {songLibrary.map((songTemplate, index) => (
              <div key={index} className="bg-white p-3 rounded border border-blue-200 hover:border-blue-400 transition-colors">
                <div className="flex justify-between items-start mb-2">
                  <div>
                    <h4 className="font-medium text-gray-800 text-sm">{songTemplate.title}</h4>
                    <div className="flex gap-2 text-xs text-gray-600">
                      <span className="bg-blue-100 px-2 py-1 rounded">
                        {getCategoryLabel(songTemplate.category)}
                      </span>
                      {songTemplate.author && <span>par {songTemplate.author}</span>}
                    </div>
                  </div>
                  <button
                    onClick={() => handleAddFromLibrary(songTemplate)}
                    className="px-2 py-1 bg-blue-600 text-white text-xs rounded hover:bg-blue-700 transition-colors"
                  >
                    Ajouter
                  </button>
                </div>
                <div className="text-xs text-gray-600 line-clamp-2">
                  {songTemplate.lyrics.split('\n')[0]}...
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {showAddForm && (
        <div className="mb-6 bg-gray-50 rounded-lg p-4">
          <h3 className="font-medium text-gray-800 mb-4">
            {editingId ? 'Modifier le chant' : 'Nouveau chant'}
          </h3>
          
          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label htmlFor="song-title" className="block text-sm font-medium text-gray-700 mb-1">
                  Titre *
                </label>
                <input
                  type="text"
                  id="song-title"
                  value={formData.title}
                  onChange={(e) => setFormData({ ...formData, title: e.target.value })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  required
                />
              </div>
              
              <div>
                <label htmlFor="song-category" className="block text-sm font-medium text-gray-700 mb-1">
                  Catégorie
                </label>
                <select
                  id="song-category"
                  value={formData.category}
                  onChange={(e) => setFormData({ ...formData, category: e.target.value as Song['category'] })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                >
                  {categories.map(category => (
                    <option key={category.value} value={category.value}>
                      {category.label}
                    </option>
                  ))}
                </select>
              </div>
              
              <div>
                <label htmlFor="song-author" className="block text-sm font-medium text-gray-700 mb-1">
                  Auteur
                </label>
                <input
                  type="text"
                  id="song-author"
                  value={formData.author}
                  onChange={(e) => setFormData({ ...formData, author: e.target.value })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>
              
              <div>
                <label htmlFor="song-melody" className="block text-sm font-medium text-gray-700 mb-1">
                  Mélodie
                </label>
                <input
                  type="text"
                  id="song-melody"
                  value={formData.melody}
                  onChange={(e) => setFormData({ ...formData, melody: e.target.value })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>
            </div>
            
            <div>
              <label htmlFor="song-lyrics" className="block text-sm font-medium text-gray-700 mb-1">
                Paroles *
              </label>
              <textarea
                id="song-lyrics"
                value={formData.lyrics}
                onChange={(e) => setFormData({ ...formData, lyrics: e.target.value })}
                rows={6}
                className="w-full px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                placeholder="Entrez les paroles du chant..."
                required
              />
            </div>
            
            <div className="flex gap-3">
              <button
                type="submit"
                className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors"
              >
                {editingId ? 'Modifier' : 'Ajouter'}
              </button>
              <button
                type="button"
                onClick={handleCancel}
                className="px-4 py-2 bg-gray-300 text-gray-700 rounded-md hover:bg-gray-400 transition-colors"
              >
                Annuler
              </button>
            </div>
          </form>
        </div>
      )}

      <div className="space-y-4">
        {songs.length === 0 ? (
          <p className="text-gray-500 text-center py-8">
            Aucun chant ajouté pour le moment.
          </p>
        ) : (
          sortedSongs.map((song, index) => (
            <div key={song.id} className="border rounded-lg p-4 bg-amber-50 border-amber-200">
              <div className="flex justify-between items-start mb-3">
                <div>
                  <div className="flex items-center gap-3">
                    <span className="bg-amber-600 text-white text-xs font-bold px-2 py-1 rounded">
                      {index + 1}
                    </span>
                    <h3 className="font-semibold text-gray-800">{song.title}</h3>
                  </div>
                  <div className="flex gap-4 text-sm text-gray-600">
                    <span className="bg-amber-200 px-2 py-1 rounded">
                      {getCategoryLabel(song.category)}
                    </span>
                    {song.author && <span>Auteur: {song.author}</span>}
                    {song.melody && <span>Mélodie: {song.melody}</span>}
                  </div>
                </div>
                
                <div className="flex gap-2">
                  <div className="flex flex-col">
                    <button
                      onClick={() => onReorderSong(song.id, 'up')}
                      disabled={index === 0}
                      className="p-1 text-gray-600 hover:bg-gray-100 rounded disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
                      title="Monter"
                    >
                      <ChevronUp className="w-3 h-3" />
                    </button>
                    <button
                      onClick={() => onReorderSong(song.id, 'down')}
                      disabled={index === sortedSongs.length - 1}
                      className="p-1 text-gray-600 hover:bg-gray-100 rounded disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
                      title="Descendre"
                    >
                      <ChevronDown className="w-3 h-3" />
                    </button>
                  </div>
                  <button
                    onClick={() => handleEdit(song)}
                    className="p-2 text-blue-600 hover:bg-blue-100 rounded-md transition-colors"
                    title="Modifier"
                  >
                    <Edit className="w-4 h-4" />
                  </button>
                  <button
                    onClick={() => onDeleteSong(song.id)}
                    className="p-2 text-red-600 hover:bg-red-100 rounded-md transition-colors"
                    title="Supprimer"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>
              
              <div className="text-gray-700 text-sm leading-relaxed whitespace-pre-line bg-white p-3 rounded border">
                {song.lyrics}
              </div>
            </div>
          ))
        )}
      </div>
    </div>
  );
}