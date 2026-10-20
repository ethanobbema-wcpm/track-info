import sys
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from metastarter_import import parse_admin_import, composers_from_cell


def xlsx(rows):
    w = Workbook()
    w.active.append(['Filename', 'CDTitle', 'TrackTitle', 'BPM', 'Key', 'Meter', 'Composer', 'Artist', None])
    for row in rows:
        w.active.append(row)
    data = BytesIO()
    w.save(data)
    return data.getvalue()


def docx(rows):
    from xml.sax.saxutils import escape
    rows = [['TRACK TITLE', 'BPM', 'KEY', 'METER', 'INSTRUMENTATION']] + rows
    xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>'
    for row in rows:
        xml += '<w:tr>' + ''.join('<w:tc><w:p><w:r><w:t>' + escape(c) + '</w:t></w:r></w:p></w:tc>' for c in row) + '</w:tr>'
    xml += '</w:tbl></w:body></w:document>'
    data = BytesIO()
    with zipfile.ZipFile(data, 'w') as z:
        z.writestr('word/document.xml', xml)
    return data.getvalue()


class ImportTests(unittest.TestCase):
    def row(self, number=1, title='One', filename=None):
        return [filename or f'ABC001_{number:03d}_{title}_Full.wav', 'Album', title, 120, 'Bm', '4/4',
                'First Writer (BMI) 99% [001234], Second Writer (ASCAP) 1% [005678]', '', '']

    def test_template_blank_rows_and_alternate_mixes_are_not_tracks(self):
        data = xlsx([['Volume_(0)xx_Track Title_Full.wav', None, 'Double-check with column A'],
                     self.row(2, 'Two'), self.row(), self.row(1, 'One', 'ABC001_001_One_30.wav'), []])
        album, notes = parse_admin_import(data, None, [], ['Bm'])
        self.assertEqual([t['track_title'] for t in album['tracks']], ['One', 'Two'])
        self.assertEqual(album['tracks'][0]['composers'][1]['split'], 1)
        self.assertEqual(album['tracks'][0]['composers'][0]['cae_ipi'], '001234')
        self.assertTrue(any('non-Full' in n for n in notes))

    def test_word_matches_title_instead_of_row_order_and_maps_instruments(self):
        album, notes = parse_admin_import(xlsx([self.row(), self.row(2, 'Two')]),
            docx([['Two', '125', 'Bm', '4/4', 'Electric Bass, Synth Keys, Alien Thing'],
                  ['One', '121', 'Bm', '3/4', 'Drumkit']]),
            ['Bass - Electric', 'Synthesizer', 'Drum Kit'], ['Bm'])
        self.assertEqual(album['tracks'][0]['bpm'], 121)
        self.assertEqual(album['tracks'][0]['instrumentation'], ['Drum Kit'])
        self.assertEqual(album['tracks'][1]['instrumentation'], ['Bass - Electric', 'Synthesizer'])
        self.assertTrue(any('Alien Thing' in n for n in notes))
        self.assertTrue(any('used Word' in n for n in notes))

    def test_single_composer_old_unnamed_dsp_column(self):
        row = self.row()
        row[6:] = ['Writer (BMI) 100% [1234]', 'Artist', '(https://music.apple.com/us/artist/123, https://open.spotify.com/artist/456)']
        album, _ = parse_admin_import(xlsx([row]), None, [], ['Bm'])
        composer = album['tracks'][0]['composers'][0]
        self.assertEqual(composer['artist_name'], 'Artist')
        self.assertEqual(composer['spotify_link'], 'https://open.spotify.com/artist/456')
        self.assertEqual(composer['apple_music_link'], 'https://music.apple.com/us/artist/123')

    def test_duplicate_and_unrelated_documents_rejected(self):
        with self.assertRaisesRegex(ValueError, 'Duplicate Full'):
            parse_admin_import(xlsx([self.row(), self.row()]), None, [], ['Bm'])
        with self.assertRaisesRegex(ValueError, 'No matching'):
            parse_admin_import(xlsx([self.row()]), docx([['Other album', '', '', '', '']]), [], ['Bm'])

    def test_missing_composer_split_remains_blank(self):
        row = self.row()
        row[6] = ''
        album, _ = parse_admin_import(xlsx([row]), None, [], ['Bm'])
        self.assertIsNone(album['tracks'][0]['composers'][0]['split'])

    def test_invalid_bpm_key_and_composer_are_not_guessed(self):
        row = self.row()
        row[3] = 'NaN'
        row[4] = 'unknown'
        row[6] = 'Unformatted Writer'
        album, notes = parse_admin_import(xlsx([row]), None, [], ['Bm'])
        track = album['tracks'][0]
        self.assertIsNone(track['bpm'])
        self.assertEqual(track['key'], '')
        self.assertIsNone(track['composers'][0]['split'])
        self.assertEqual(track['producer'], '')
        self.assertGreaterEqual(len(notes), 3)


if __name__ == '__main__':
    unittest.main()
