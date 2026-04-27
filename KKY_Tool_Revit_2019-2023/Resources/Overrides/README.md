# Review Export Overrides

Recommended workflow:

1. Edit `docs/review_export_feature_sheet_matrix.xlsx`.
2. Run:
   `python tools/sync_review_export_feature_sheet_matrix_to_overrides.py`
3. Build the project.

Direct edit workflow is also possible:

1. Edit `review_export_message_overrides.xlsx` in this folder.
2. Build the project.

Current runtime mapping:

- Korean export:
  - header = `한글 헤더 override` if set, otherwise `원본/내부 열명`
  - text = `한글 override` if set, otherwise `원본(raw)`
- English export:
  - header = `영문 헤더 override` if set, otherwise `영문 export 열명`
  - text = `영문 override` if set, otherwise the existing English export text

`docs/review_export_feature_sheet_matrix.xlsx` is the easier editing surface because it is organized by:

- feature
- actual export section/sheet
- locale
- condition
- final visible headers

`review_export_message_overrides.xlsx` is the runtime/source workbook that is copied into build output and read by the add-in.

Notes:

- `Resources\**\*` is copied during build, so edits synced into this folder are included automatically.
- The workbook path takes precedence over JSON.
- `review_export_message_overrides.json` remains only as a fallback path when needed.
