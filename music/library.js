import { albums } from './catalog.js'

const musicTracks = albums.flatMap(album => album.tracks.map(track => ({
  id: track.id,
  title: track.title,
  artist: track.artist,
  album: album.name,
  genre: 'Soundtrack',
  duration: track.duration,
  addedAt: album.addedAt,
  src: new URL('./audio/' + track.id + '.mp3', import.meta.url).href,
  artwork: new URL('./' + album.artwork, import.meta.url).href,
})))

const musicPlaylists = albums.map(album => ({
  id: album.id,
  name: album.name,
  trackIds: album.tracks.map(track => track.id),
}))

function appendNew(items, additions) {
  const existing = new Set(items.map(item => item.id))
  return [...items, ...additions.filter(item => !existing.has(item.id))]
}

export function mergeMusicLibrary(state) {
  const installedIds = new Set(state.tracks.map(track => track.id))
  if (musicTracks.every(track => installedIds.has(track.id))) return null
  return {
    tracks: appendNew(state.tracks, musicTracks),
    playlists: appendNew(state.playlists, musicPlaylists),
  }
}
