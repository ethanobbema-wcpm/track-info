# Track Info

Track Info is a Streamlit app for collecting album, track, instrumentation, and composer metadata, then exporting the result as an Excel workbook with one row per track.

The app also supports a workbook round-trip workflow through the `Need to fix Track Info?` section at the bottom of the page, where users can upload a previously exported Track Info workbook, repopulate the form, edit it, and export a fresh version.

All data-entry fields are required before export except `DSP Links (Spotify, Apple Music)` and `Artist Name (Optional)`. Missing or invalid required fields appear in the collapsed `Required Fields` section.

## Files

- `streamlit_app.py` is the Streamlit entrypoint.
- `requirements.txt` lists the Python packages needed by Streamlit Community Cloud.
- `taxonomy_instruments.xlsx` is the bundled instrument taxonomy used by the app.
- `vocal_sublist.xlsx` is the bundled vocal sub-list shown when `Vocals` is selected.
- `key_selection.xlsx` is the bundled key dropdown list used by the app.

## Run Locally

```bash
cd track_info_streamlit
python3 -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy On Streamlit Community Cloud

Deploy this repository and set the app entrypoint to:

```text
track_info_streamlit/streamlit_app.py
```

The dependency file is in the same directory as the entrypoint, which Streamlit Community Cloud supports.

## Instrument Taxonomy

The app loads instrument names from `taxonomy_instruments.xlsx`, removes duplicates, and uses those names as the searchable options for each track's Instrumentation field.

When `Vocals` is selected in a track's Instrumentation field, the app shows a `Vocal Sub-list` multi-select loaded from `vocal_sublist.xlsx`.

The `Key` field uses the bundled key list, includes a `No Key` option, and supports `Multi-Key?` mode that exports selected keys as a comma-separated value in the `Key` column.

## Excel Export

The exported workbook has a `Track Info` sheet. Each track appears on its own row, with album and WCPM producer fields repeated on each row and composer columns expanded as needed. `Vocal Sub-list` is always included as an export column and remains blank for tracks without `Vocals`. `CAE/IPI` is required for each composer. Composer splits are exported as Excel percentage cells:

- `Composer 1`, `PRO Affiliation 1`, `CAE/IPI 1`, `Splits 1 (%)`, `DSP Links (Spotify, Apple Music) 1`, `Artist Name 1 (Optional)`
- `Composer 2`, `PRO Affiliation 2`, `CAE/IPI 2`, `Splits 2 (%)`, `DSP Links (Spotify, Apple Music) 2`, `Artist Name 2 (Optional)`
- Additional composer groups when needed

Previously exported Track Info workbooks can be uploaded back into the app, and the form state will be reconstructed from the workbook so users can continue editing and export again.
