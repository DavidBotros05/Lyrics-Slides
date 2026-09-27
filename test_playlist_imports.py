import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import Create_Lyrics as core
import app


APPLE_PLAYLIST_LINK = (
    'https://music.apple.com/ca/playlist/sunday-set/pl.1234-abcd?l=en'
)
SPOTIFY_PLAYLIST_LINK = (
    'https://open.spotify.com/playlist/37i9dQZF1DXcBWIGoYBM5M'
)
APPLE_PLAYLIST_HTML = '''
<html><head>
  <script type="application/ld+json">
    {
      "@context": "https://schema.org",
      "@type": "MusicPlaylist",
      "name": "Sunday Set",
      "numTracks": 2,
      "track": [
        {
          "@type": "MusicRecording",
          "name": "Grace",
          "byArtist": {"@type": "MusicGroup", "name": "Artist One"}
        },
        {
          "@type": "MusicRecording",
          "name": "Hope",
          "byArtist": {"@type": "MusicGroup", "name": "Artist Two"}
        }
      ]
    }
  </script>
</head></html>
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

    def test_detects_supported_playlist_service_from_the_link(self):
        self.assertEqual('apple_music', core.detect_playlist_service(
            APPLE_PLAYLIST_LINK))
        self.assertEqual('spotify', core.detect_playlist_service(
            SPOTIFY_PLAYLIST_LINK))

    def test_parses_public_apple_music_playlist_link(self):
        self.assertEqual(
            ('ca', 'pl.1234-abcd'),
            core._apple_music_parse_playlist_link(APPLE_PLAYLIST_LINK),
        )

    def test_reads_tracks_and_artists_from_public_playlist_page(self):
        response = Mock(text=APPLE_PLAYLIST_HTML)
        response.raise_for_status = Mock()
        with patch.object(core.requests, 'get', return_value=response) as get:
            result = core.fetch_apple_music_playlist(APPLE_PLAYLIST_LINK)

        self.assertEqual('Sunday Set', result['name'])
        self.assertEqual(
            [
                {'title': 'Grace', 'artist': 'Artist One'},
                {'title': 'Hope', 'artist': 'Artist Two'},
            ],
            result['tracks'],
        )
        self.assertEqual(2, result['total'])
        self.assertEqual('', result['note'])
        get.assert_called_once_with(
            APPLE_PLAYLIST_LINK, headers=core.BROWSER_HEADERS, timeout=20)

    def test_uses_developer_token_for_complete_apple_music_playlist(self):
        playlist_response = Mock()
        playlist_response.raise_for_status = Mock()
        playlist_response.json.return_value = {
            'data': [{'attributes': {'name': 'Sunday Set'}}],
        }
        tracks_response = Mock()
        tracks_response.raise_for_status = Mock()
        tracks_response.json.return_value = {
            'data': [
                {'attributes': {'name': 'Grace', 'artistName': 'Artist One'}},
                {'attributes': {'name': 'Hope', 'artistName': 'Artist Two'}},
            ],
            'next': None,
        }
        with (
            patch.object(core, 'APPLE_MUSIC_DEVELOPER_TOKEN', 'test-token'),
            patch.object(
                core.requests, 'get',
                side_effect=[playlist_response, tracks_response],
            ) as get,
        ):
            result = core.fetch_apple_music_playlist(APPLE_PLAYLIST_LINK)

        self.assertEqual('Sunday Set', result['name'])
        self.assertEqual(2, result['total'])
        self.assertEqual('', result['note'])
        self.assertEqual(2, get.call_count)
        self.assertEqual(
            'Bearer test-token',
            get.call_args_list[0].kwargs['headers']['Authorization'],
        )

    def test_rejects_non_playlist_apple_music_link(self):
        self.assertIsNone(core.fetch_apple_music_playlist(
            'https://music.apple.com/ca/album/example/123456789'))
        self.assertIn('Apple Music playlist', core.LAST_ERROR)

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

    def test_reads_artist_from_apple_music_subtitle(self):
        item = {'title': 'Hope', 'subtitle': 'Artist Two'}
        self.assertEqual(
            {'title': 'Hope', 'artist': 'Artist Two'},
            core._apple_music_track(item),
        )

    def test_public_page_fallback_keeps_nested_artist_relationship(self):
        html = '''
        <meta property="og:title" content="Sunday Set">
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

    def test_api_returns_importable_tracks_for_apple_music_playlist(self):
        handler = object.__new__(app.Handler)
        payloads = []
        handler._send_json = lambda payload, status=200: payloads.append(payload)
        playlist = {
            'name': 'Sunday Set',
            'tracks': [{'title': 'Grace', 'artist': 'Artist One'}],
            'total': 1,
            'note': '',
        }

        with patch.object(core, 'fetch_apple_music_playlist', return_value=playlist):
            handler.handle_playlist({'link': APPLE_PLAYLIST_LINK})

        self.assertEqual({'ok': True, **playlist}, payloads[0])

    def test_cloud_ui_has_one_auto_detecting_playlist_import(self):
        html = Path(__file__).with_name('ui.html').read_text(encoding='utf-8')
        self.assertEqual(1, html.count('id="playlistLink"'))
        self.assertIn("api('/api/playlist', {link})", html)
        self.assertNotIn('id="spotifyLink"', html)
        self.assertNotIn('id="appleMusicLink"', html)


if __name__ == '__main__':
    unittest.main()
