"""Session-scoped, server-verified access to legacy metadata importing."""
from __future__ import annotations

import hashlib
import hmac
import os
import time

import streamlit as st

from metastarter_import import MAX_FILE_BYTES, parse_admin_import

AUTH_SECONDS = 30 * 60
# Server-side verifier for the owner-selected import password. No plaintext password is bundled.
DEFAULT_ADMIN_PASSWORD_HASH = 'pbkdf2_sha256$600000$7b4ae2bd762c86eb1bcdbca7448569c0$ed9f21c00d0a536ac29fe81f5edb9be2f76106db19570057c3b36905963accfb'


def admin_password() -> str:
    password = os.environ.get('TRACK_INFO_ADMIN_PASSWORD', '')
    if not password:
        try:
            password = str(st.secrets.get('TRACK_INFO_ADMIN_PASSWORD', ''))
        except (FileNotFoundError, st.errors.StreamlitSecretNotFoundError):
            password = ''
    return password or DEFAULT_ADMIN_PASSWORD_HASH


def verify_password(entered: str, credential: str) -> bool:
    if credential == DEFAULT_ADMIN_PASSWORD_HASH:
        _, iterations, salt, expected = credential.split('$')
        actual = hashlib.pbkdf2_hmac('sha256', entered.encode(), bytes.fromhex(salt), int(iterations)).hex()
        return hmac.compare_digest(actual, expected)
    return hmac.compare_digest(entered.encode(), credential.encode())


def authorized(password: str) -> bool:
    return bool(password) and (
        st.session_state.get('_admin_auth_fingerprint') == hashlib.sha256(password.encode()).hexdigest()
        and time.time() < st.session_state.get('_admin_auth_until', 0)
    )


def lock_admin() -> None:
    for key in list(st.session_state):
        if key.startswith('_admin_') and key not in ('_admin_attempts', '_admin_retry_after'):
            del st.session_state[key]


def render_admin_import(instruments: list[str], keys: list[str]) -> None:
    if st.button('Private import', key='_admin_toggle', help='Password-protected metadata import'):
        st.session_state['_admin_panel_open'] = not st.session_state.get('_admin_panel_open', False)
    if not st.session_state.get('_admin_panel_open'):
        return
    with st.expander('Private metadata import', expanded=True):
        password = admin_password()
        if not password:
            st.info('Set TRACK_INFO_ADMIN_PASSWORD in Streamlit secrets to enable private importing.')
            return
        if not authorized(password):
            # Clear any expired upload/preview before presenting the login form.
            for key in ('_admin_preview', '_admin_xlsx', '_admin_docx'):
                st.session_state.pop(key, None)
            with st.form('_admin_login'):
                entered = st.text_input('Import password', type='password', key='_admin_password')
                submitted = st.form_submit_button('Unlock import')
            if submitted:
                st.session_state.pop('_admin_password', None)
                now = time.time()
                if now < st.session_state.get('_admin_retry_after', 0):
                    st.error('Too many attempts. Wait a minute before trying again.')
                elif verify_password(entered, password):
                    st.session_state['_admin_auth_fingerprint'] = hashlib.sha256(password.encode()).hexdigest()
                    st.session_state['_admin_auth_until'] = now + AUTH_SECONDS
                    st.session_state['_admin_attempts'] = 0
                    st.rerun()
                else:
                    attempts = st.session_state.get('_admin_attempts', 0) + 1
                    st.session_state['_admin_attempts'] = attempts
                    if attempts >= 5:
                        st.session_state['_admin_retry_after'] = now + 60
                        st.session_state['_admin_attempts'] = 0
                    st.error('Incorrect password.')
            return
        st.caption('Upload one completed Metastarter, with an optional matching Track Info Word document. Files are processed in this session. Access expires after 30 minutes.')
        if st.button('Lock private import', key='_admin_lock'):
            lock_admin()
            st.rerun()
        workbook = st.file_uploader('Completed Metastarter (.xlsx)', type=['xlsx'], key='_admin_xlsx')
        document = st.file_uploader('Track Info (.docx, optional)', type=['docx'], key='_admin_docx')
        signature = tuple((f.name, hashlib.sha256(f.getvalue()).hexdigest()) if f is not None else None
                          for f in (workbook, document))
        if st.session_state.get('_admin_file_signature') != signature:
            st.session_state.pop('_admin_preview', None)
            st.session_state['_admin_file_signature'] = signature
        if st.button('Preview import', disabled=workbook is None, key='_admin_parse'):
            # Recheck at the action boundary, including session expiry.
            if not authorized(admin_password()):
                lock_admin()
                st.rerun()
            st.session_state.pop('_admin_preview', None)
            try:
                if any(f is not None and f.size > MAX_FILE_BYTES for f in (workbook, document)):
                    raise ValueError('Each import file must be 20 MB or smaller.')
                album, warnings = parse_admin_import(workbook.getvalue(), document.getvalue() if document else None, instruments, keys)
                st.session_state['_admin_preview'] = album, warnings
            except Exception as exc:
                st.error(f'Import could not be read: {exc}')
        preview = st.session_state.get('_admin_preview')
        if preview:
            album, warnings = preview
            st.write(f"{album['album_name'] or 'Untitled album'} · {len(album['tracks'])} tracks")
            st.dataframe([{'Track': t['track_number'], 'Title': t['track_title'], 'BPM': t['bpm'],
                           'Key': t['key'], 'Meter': t['meter'], 'Instrumentation': ', '.join(t['instrumentation']),
                           'Composers': ', '.join(c['composer'] for c in t['composers'])}
                          for t in album['tracks']], hide_index=True)
            if warnings:
                with st.expander('Import notes', expanded=True):
                    for warning in warnings:
                        st.warning(warning)
            st.caption('Loading replaces the current form. Missing information remains blank. Review Required Fields before exporting.')
            if st.button('Replace form with this import', key='_admin_apply', type='primary'):
                if not authorized(admin_password()):
                    lock_admin()
                    st.rerun()
                st.session_state['_pending_track_info_import'] = album
                st.session_state['_track_info_import_message'] = f"Loaded {len(album['tracks'])} tracks from the private import. Review Required Fields before exporting."
                st.session_state['_admin_last_notes'] = warnings
                st.session_state.pop('_admin_preview', None)
                st.rerun()
        if st.session_state.get('_admin_last_notes'):
            with st.expander('Last import notes'):
                for warning in st.session_state['_admin_last_notes']:
                    st.warning(warning)
