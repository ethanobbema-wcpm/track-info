import hashlib
import os
from pathlib import Path
import sys
import time
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from streamlit.testing.v1 import AppTest
from test_metastarter_import import xlsx

APP = str(Path(__file__).resolve().parents[1] / 'streamlit_app.py')
SUBMIT = 'FormSubmitter:_admin_login-Unlock import'
PASSWORD = 'test-password-💿'


class Upload:
    name = 'test.xlsx'

    def __init__(self, data):
        self.data = data
        self.size = len(data)

    def getvalue(self):
        return self.data


class AdminTests(unittest.TestCase):
    def setUp(self):
        self.env = patch.dict(os.environ, {'TRACK_INFO_ADMIN_PASSWORD': PASSWORD})
        self.env.start()
        self.addCleanup(self.env.stop)

    def open_login(self):
        app = AppTest.from_file(APP, default_timeout=30).run()
        self.assertFalse(app.exception)
        self.assertTrue(any(b.key == '_admin_toggle' for b in app.button))
        self.assertEqual(len(app.get('file_uploader')), 0)
        self.assertNotIn('admin', app.query_params)
        app.button(key='_admin_toggle').click().run()
        return app

    def unlock(self, app):
        app.text_input(key='_admin_password').input(PASSWORD)
        app.button(key=SUBMIT).click().run()
        self.assertFalse(app.exception)

    def test_password_gate_expiry_rotation_and_lock(self):
        app = self.open_login()
        self.assertEqual(len(app.get('file_uploader')), 0)
        app.text_input(key='_admin_password').input('wrong')
        app.button(key=SUBMIT).click().run()
        self.assertTrue(any(e.value == 'Incorrect password.' for e in app.error))
        self.assertEqual(len(app.get('file_uploader')), 0)
        self.unlock(app)
        self.assertEqual(len(app.get('file_uploader')), 2)
        app.session_state['_admin_auth_until'] = 0
        app.run()
        self.assertEqual(len(app.get('file_uploader')), 0)
        self.unlock(app)
        with patch.dict(os.environ, {'TRACK_INFO_ADMIN_PASSWORD': 'rotated-password'}):
            app.run()
            self.assertEqual(len(app.get('file_uploader')), 0)
        self.unlock(app)
        app.button(key='_admin_lock').click().run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.get('file_uploader')), 0)

    def test_bundled_password_hash_and_override(self):
        from admin_import_ui import admin_password, verify_password, DEFAULT_ADMIN_PASSWORD_HASH
        configured = 'test-import-password'
        salt = '00' * 16
        digest = hashlib.pbkdf2_hmac('sha256', configured.encode(), bytes.fromhex(salt), 600000).hex()
        verifier = f'pbkdf2_sha256$600000${salt}${digest}'
        with patch.dict(os.environ, {'TRACK_INFO_ADMIN_PASSWORD': ''}), patch('streamlit.secrets', {}), patch('admin_import_ui.DEFAULT_ADMIN_PASSWORD_HASH', verifier):
            self.assertEqual(admin_password(), verifier)
            self.assertTrue(verify_password(configured, admin_password()))
            self.assertFalse(verify_password('wrong', admin_password()))
            app = self.open_login()
            app.text_input(key='_admin_password').input(configured)
            app.button(key=SUBMIT).click().run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.get('file_uploader')), 2)
        self.assertEqual(admin_password(), PASSWORD)
        self.assertFalse(verify_password(configured, admin_password()))

    def test_cooldown_rejects_even_correct_password(self):
        app = self.open_login()
        for _ in range(5):
            app.text_input(key='_admin_password').input('wrong')
            app.button(key=SUBMIT).click().run()
        app.text_input(key='_admin_password').input(PASSWORD)
        app.button(key=SUBMIT).click().run()
        self.assertEqual(len(app.get('file_uploader')), 0)
        self.assertTrue(any('Too many attempts' in e.value for e in app.error))

    def test_preview_does_not_replace_until_apply_and_missing_split_stays_blank(self):
        upload = Upload(xlsx([['ABC001_001_One_Full.wav', 'Imported album', 'One', 120,
                              'Bm', '4/4', 'Unknown Writer', '', '']]))
        app = self.open_login()
        self.unlock(app)
        app.text_input(key='album_name').input('Existing album').run()
        # AppTest does not expose upload interactions. Substitute only the file picker,
        # then exercise the real preview, parser, apply button and form rerun.
        def picker(label, **kwargs):
            return upload if kwargs['key'] == '_admin_xlsx' else None
        with patch('streamlit.file_uploader', side_effect=picker):
            app.run()
            app.button(key='_admin_parse').click().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.session_state['album_name'], 'Existing album')
            self.assertEqual(len(app.dataframe), 1)
            app.button(key='_admin_apply').click().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.session_state['album_name'], 'Imported album')
            self.assertEqual(app.session_state['track_1_title'], 'One')
            self.assertEqual(app.session_state['track_1_composer_1_split'], '')
            self.assertTrue(any('Splits must be a number' in w.value for w in app.warning))
        # Locking the private area preserves the imported form on the regular page.
        app.button(key='_admin_lock').click().run()
        self.assertFalse(app.exception)
        self.assertEqual(app.session_state['track_1_title'], 'One')
        self.assertEqual(len(app.get('file_uploader')), 0)


if __name__ == '__main__':
    unittest.main()
