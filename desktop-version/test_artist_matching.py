import unittest
import os
import tempfile
from unittest.mock import patch

import Create_Lyrics as core
import app as desktop_app


def candidate(path, artist):
    return {
        'path': path,
        'file': path,
        'title': 'Test Song',
        'artist': artist,
        'artist_key': core._song_lookup_key(artist),
    }


class ArtistLibraryMatchingTests(unittest.TestCase):
    def resolve(self, requested_artist, candidates):
        with (
            patch.object(core, '_library_candidates', return_value=candidates),
            patch.object(
                core,
                'place_existing_song_in_output',
                side_effect=lambda path, *_: {'path': path},
            ),
        ):
            return core.resolve_existing_song_from_all_songs(
                requested_artist, 'Test Song', '', '/tmp/output')

    def test_single_artist_matches_saved_multi_artist_credit(self):
        result = self.resolve('Artist A', [candidate('saved.pptx', 'Artist A & Artist B')])
        self.assertEqual('placed', result['status'])
        self.assertEqual('saved.pptx', result['result']['path'])

    def test_equivalent_multi_artist_separators_match(self):
        result = self.resolve(
            'Artist A, Artist B',
            [candidate('saved.pptx', 'Artist A feat. Artist B')],
        )
        self.assertEqual('placed', result['status'])

    def test_exact_match_wins_over_partial_multi_artist_match(self):
        result = self.resolve(
            'Artist A',
            [
                candidate('collaboration.pptx', 'Artist A & Artist B'),
                candidate('exact.pptx', 'Artist A'),
            ],
        )
        self.assertEqual('exact.pptx', result['result']['path'])

    def test_two_partial_matches_still_require_a_choice(self):
        result = self.resolve(
            'Artist A',
            [
                candidate('one.pptx', 'Artist A & Artist B'),
                candidate('two.pptx', 'Artist A & Artist C'),
            ],
        )
        self.assertEqual('ambiguous', result['status'])

    def test_library_save_prefers_exact_credit_over_partial_match(self):
        candidates = [
            candidate('collaboration.pptx', 'Artist A & Artist B'),
            candidate('exact.pptx', 'Artist A'),
        ]
        with patch.object(core, '_library_candidates', return_value=candidates):
            destination = core._library_destination_for_song('Test Song', 'Artist A')
        self.assertEqual('exact.pptx', destination)


class OutputFolderReuseTests(unittest.TestCase):
    def test_chosen_folder_song_is_used_without_copying_or_moving(self):
        with tempfile.TemporaryDirectory() as output_dir:
            existing_path = os.path.join(output_dir, 'Test Song.pptx')
            with open(existing_path, 'wb'):
                pass

            with patch.object(
                core, '_presentation_meta', return_value=('Test Song', 'Artist A')
            ):
                result = core.resolve_existing_song_from_output(
                    'Artist A', 'Test Song', '', output_dir)

            self.assertEqual('placed', result['status'])
            self.assertEqual(existing_path, result['result']['path'])
            self.assertEqual('Chosen output folder', result['result']['source'])
            self.assertTrue(os.path.exists(existing_path))

    def test_cleanup_preserves_matches_for_every_song_in_batch(self):
        with tempfile.TemporaryDirectory() as output_dir:
            first = os.path.join(output_dir, 'First Song.pptx')
            second = os.path.join(output_dir, 'Second Song.pptx')
            unrelated = os.path.join(output_dir, 'Old Song.pptx')
            for path in (first, second, unrelated):
                with open(path, 'wb'):
                    pass

            protected = core.existing_output_paths_for_songs(output_dir, [
                {'title': 'First Song', 'artist': 'Artist A'},
                {'title': 'Second Song', 'artist': 'Artist B'},
            ])
            changed = core.prepare_selected_output_folder(
                output_dir, core.PREVIOUS_POWERPOINTS_DELETE, protected)

            self.assertEqual(1, changed)
            self.assertTrue(os.path.exists(first))
            self.assertTrue(os.path.exists(second))
            self.assertFalse(os.path.exists(unrelated))

    def test_multiple_output_versions_require_a_choice(self):
        with tempfile.TemporaryDirectory() as output_dir:
            paths = [
                os.path.join(output_dir, 'Test Song.pptx'),
                os.path.join(output_dir, 'Test Song 2.pptx'),
            ]
            for path in paths:
                with open(path, 'wb'):
                    pass

            artists = {
                paths[0]: ('Test Song', 'Artist A'),
                paths[1]: ('Test Song', 'Artist B'),
            }
            with patch.object(
                core, '_presentation_meta', side_effect=lambda path, *_: artists[path]
            ):
                result = core.resolve_existing_song_from_output(
                    '', 'Test Song', '', output_dir)

            self.assertEqual('ambiguous', result['status'])
            self.assertEqual(2, len(result['choices']))

    def test_create_endpoint_returns_output_song_before_library_or_fetch(self):
        with tempfile.TemporaryDirectory() as output_dir:
            existing_path = os.path.join(output_dir, 'Test Song.pptx')
            result = {
                'path': existing_path,
                'file': 'Test Song.pptx',
                'title': 'Test Song',
                'artist': 'Artist A',
                'source': 'Chosen output folder',
                'delivery': 'already_present',
            }
            responses = []
            handler = object.__new__(desktop_app.Handler)
            handler._send_json = lambda payload, status=200: responses.append(payload)

            with (
                patch.object(
                    core, 'existing_output_paths_for_songs',
                    return_value=[existing_path],
                ),
                patch.object(
                    core, 'resolve_existing_song_from_output',
                    return_value={'status': 'placed', 'result': result},
                ),
                patch.object(core, 'prepare_selected_output_folder') as prepare,
                patch.object(core, 'resolve_existing_song_from_all_songs') as all_songs,
                patch.object(desktop_app, 'fetch_song_for_request') as fetch,
            ):
                handler.handle_create_one({
                    'artist': 'Artist A',
                    'title': 'Test Song',
                    'output_folder': output_dir,
                    'previous_powerpoints_action': core.PREVIOUS_POWERPOINTS_DELETE,
                    'preserve_songs': [{'artist': 'Artist A', 'title': 'Test Song'}],
                })

            prepare.assert_called_once()
            self.assertIn(existing_path, prepare.call_args.args[2])
            all_songs.assert_not_called()
            fetch.assert_not_called()
            self.assertEqual('Chosen output folder', responses[0]['source'])


if __name__ == '__main__':
    unittest.main()
