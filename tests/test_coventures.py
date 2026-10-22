from io import BytesIO
from pathlib import Path
import sys
import unittest

from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from streamlit_app import (COVENTURES_KEY_PATH, load_coventures_options, build_import_state,
                          build_excel_workbook, parse_imported_workbook, validation_messages)
from metastarter_import import parse_admin_import
from test_metastarter_import import xlsx

APP = str(Path(__file__).resolve().parents[1] / 'streamlit_app.py')


class CoVenturesTests(unittest.TestCase):
    def test_toggle_autofills_producer_and_restores_previous_value(self):
        app = AppTest.from_file(APP, default_timeout=30).run()
        self.assertFalse(app.exception)
        self.assertFalse(any(s.key == 'album_category' for s in app.selectbox))
        app.text_input(key='producer').input('Original Producer').run()
        app.checkbox(key='coventures').check().run()
        self.assertFalse(app.exception)
        self.assertEqual(app.text_input(key='producer').value, 'None')
        self.assertTrue(app.text_input(key='producer').disabled)
        categories, sub = load_coventures_options(COVENTURES_KEY_PATH.stat().st_mtime_ns)
        self.assertEqual(app.selectbox(key='album_category').options, categories)
        self.assertEqual(app.multiselect(key='album_sub_categories').options, sub)
        self.assertTrue(any('choose one Category' in m.value for m in app.warning))
        self.assertTrue(any('exactly three Sub-Categories' in m.value for m in app.warning))
        app.selectbox(key='album_category').select(categories[0]).run()
        app.multiselect(key='album_sub_categories').set_value(sub[:2]).run()
        self.assertTrue(any('exactly three Sub-Categories' in m.value for m in app.warning))
        app.multiselect(key='album_sub_categories').set_value(sub[:3]).run()
        self.assertFalse(any('CoVentures:' in m.value for m in app.warning))
        self.assertEqual(app.text_area(key='album_description').value, '')
        app.checkbox(key='coventures').uncheck().run()
        self.assertFalse(app.exception)
        self.assertEqual(app.text_input(key='producer').value, 'Original Producer')
        self.assertFalse(app.text_input(key='producer').disabled)
        self.assertFalse(any(s.key == 'album_category' for s in app.selectbox))
        self.assertFalse(any('CoVentures:' in m.value for m in app.warning))

    def test_excel_round_trip_validation_and_legacy_import(self):
        categories, sub = load_coventures_options(COVENTURES_KEY_PATH.stat().st_mtime_ns)
        album, _ = parse_admin_import(xlsx([['ABC001_001_One_Full.wav', 'Album', 'One', 120,
            'Bm', '4/4', 'Writer (BMI) 100% [12345]', '', '']]), None, [], ['Bm'])
        track = album['tracks'][0]
        track.update(producer='None', instrumentation='Piano', vocal_sublist='', featured_instrument='Piano',
                     meter_numerator='4', meter_denominator='4', coventures=True,
                     category=categories[0], sub_categories=sub[:3], album_description='Line one\nLine two')
        self.assertEqual(validation_messages([track]), [])
        for selection in [[], sub[:2], sub[:4], [sub[0]] * 3]:
            track['sub_categories'] = selection
            self.assertTrue(any('exactly three' in m for m in validation_messages([track])))
        track['sub_categories'] = sub[:3]
        result = parse_imported_workbook(build_excel_workbook([track]))
        state = build_import_state(result)
        self.assertTrue(state['coventures'])
        self.assertEqual(state['producer'], 'None')
        self.assertEqual(state['album_category'], categories[0])
        self.assertEqual(state['album_sub_categories'], sub[:3])
        self.assertEqual(state['album_description'], 'Line one\nLine two')
        app = AppTest.from_file(APP, default_timeout=30)
        app.session_state['_pending_track_info_import'] = result
        app.run()
        self.assertFalse(app.exception)
        self.assertEqual(app.multiselect(key='album_sub_categories').value, sub[:3])
        track['coventures'] = False
        old_export = build_excel_workbook([track])
        w = load_workbook(BytesIO(old_export))
        self.assertNotIn('CoVentures', [c.value for c in w.active[1]])
        w.close()
        app.session_state['_pending_track_info_import'] = parse_imported_workbook(old_export)
        app.run()
        self.assertFalse(app.exception)
        self.assertFalse(app.checkbox(key='coventures').value)
        self.assertFalse(any(s.key == 'album_category' for s in app.selectbox))

    def test_reset_clears_coventures(self):
        app = AppTest.from_file(APP, default_timeout=30).run()
        app.checkbox(key='coventures').check().run()
        app.button(key='reset_track_info').click().run()
        self.assertFalse(app.exception)
        self.assertFalse(app.checkbox(key='coventures').value)
        self.assertEqual(app.text_input(key='producer').value, '')


if __name__ == '__main__':
    unittest.main()
