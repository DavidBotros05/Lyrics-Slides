import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import Create_Lyrics as core
import app as desktop_app


APPLE_PLAYLIST_LINK = (
    'https://music.apple.com/ca/playlist/sunday-set/pl.1234-abcd?l=en'
)
APPLE_PLAYLIST_HTML = '''
<script type="application/ld+json">
  {
    "@type": "MusicPlaylist", "name": "Sunday Set", "numTracks": 2,
    "track": [
      {"@type": "MusicRecording", "name": "Grace",
       "byArtist": {"name": "Artist One"}},
      {"@type": "MusicRecording", "name": "Hope",
       "byArtist": {"name": "Artist Two"}}
    ]
  }
</script>
'''


class AppleMusicPlaylistTests(unittest.TestCase):
    def test_title_only_metadata_is_enriched_by_song_id(self):
        # Apple's live page separates title-only schema from full song records.
        html = '''
        <script type="application/ld+json">
        {"@type":"MusicPlaylist","name":"Set","track":[
          {"name":"Same title","url":"https://music.apple.com/us/song/a/123"},
          {"name":"Same title","url":"https://music.apple.com/us/song/b/456"}]}
        </script>
        <script type="application/json" id="serialized-server-data">
        {"items":[
          {"title":"Same title","artistName":"Artist B",
           "contentDescriptor":{"kind":"song","identifiers":{"storeAdamID":"456"}}},
          {"title":"Same title","subtitleLinks":[{"title":"Artist A"}],
           "contentDescriptor":{"kind":"song","identifiers":{"storeAdamID":"123"}}}]}
        </script>
        '''
        with patch.object(core.requests, 'get', return_value=Mock(text=html)):
            result = core._apple_music_from_page(APPLE_PLAYLIST_LINK)
        self.assertEqual([
            {'title': 'Same title', 'artist': 'Artist A'},
            {'title': 'Same title', 'artist': 'Artist B'},
        ], result['tracks'])

    def test_reads_artist_from_nested_apple_music_relationship(self):
        item = {
            'attributes': {'name': 'Grace'},
            'relationships': {
                'artists': {
                    'data': [{'attributes': {'name': 'Artist One'}}],
                },
            },
        }
        self.assertEqual(
            {'title': 'Grace', 'artist': 'Artist One'},
            core._apple_music_track(item),
        )

    def test_public_page_fallback_keeps_nested_artist_relationship(self):
        html = '''
        <script type="application/json">
          {"data": [{"type": "songs", "attributes": {"name": "Grace"},
            "relationships": {"artists": {"data": [
              {"attributes": {"name": "Artist One"}}
            ]}}}]}
        </script>
        '''
        response = Mock(text=html)
        response.raise_for_status = Mock()
        with patch.object(core.requests, 'get', return_value=response):
            result = core.fetch_apple_music_playlist(APPLE_PLAYLIST_LINK)

        self.assertEqual(
            [{'title': 'Grace', 'artist': 'Artist One'}],
            result['tracks'],
        )

    def test_parses_and_reads_a_public_apple_music_playlist(self):
        response = Mock(text=APPLE_PLAYLIST_HTML)
        response.raise_for_status = Mock()
        with patch.object(core.requests, 'get', return_value=response):
            result = core.fetch_apple_music_playlist(APPLE_PLAYLIST_LINK)

        self.assertEqual('Sunday Set', result['name'])
        self.assertEqual([
            {'title': 'Grace', 'artist': 'Artist One'},
            {'title': 'Hope', 'artist': 'Artist Two'},
        ], result['tracks'])

    def test_rejects_an_apple_music_link_that_is_not_a_playlist(self):
        self.assertIsNone(core.fetch_apple_music_playlist(
            'https://music.apple.com/ca/album/example/123456789'))
        self.assertIn('Apple Music playlist', core.LAST_ERROR)

    def test_endpoint_detects_and_saves_last_apple_music_playlist(self):
        handler = object.__new__(desktop_app.Handler)
        payloads = []
        handler._send_json = lambda payload, status=200: payloads.append(payload)
        playlist = {
            'name': 'Sunday Set',
            'tracks': [{'title': 'Grace', 'artist': 'Artist One'}],
            'total': 1,
            'note': '',
        }
        settings = {}
        with (
            patch.object(core, 'fetch_apple_music_playlist', return_value=playlist),
            patch.object(core, 'load_app_settings', return_value=settings),
            patch.object(core, 'save_app_settings') as save,
        ):
            handler.handle_playlist({'link': APPLE_PLAYLIST_LINK})

        self.assertEqual(APPLE_PLAYLIST_LINK, settings['last_playlist'])
        save.assert_called_once_with(settings)
        self.assertEqual({'ok': True, **playlist}, payloads[0])

    def test_desktop_ui_has_one_auto_detecting_playlist_import(self):
        html = Path(__file__).with_name('ui.html').read_text(encoding='utf-8')
        self.assertEqual(1, html.count('id="playlistLink"'))
        self.assertIn("api('/api/playlist', {link})", html)
        self.assertNotIn('id="spotifyLink"', html)
        self.assertNotIn('id="appleMusicLink"', html)


if __name__ == '__main__':
    unittest.main()
