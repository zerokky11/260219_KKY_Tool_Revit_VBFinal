# GitHub Pages Update Files

This folder contains the minimum files needed for a GitHub Pages based update test.

Files:

- `index.html`
- `.nojekyll`
- `latest.json`
- `family-browser/bootstrap-index.json`
- `family-browser/bootstrap.json`

Before uploading:

1. Replace `your-github-username` in `latest.json`
2. Replace `your-repo-name` in `latest.json`
3. Upload your installer file next to `latest.json`
4. If needed, change the version number and installer filename
5. Edit the Family Browser Path tab on `index.html`, then save or upload `family-browser/bootstrap-index.json` and the selected profile JSON file

Profile JSON files can include fallback paths. Use one path per candidate; Family Browser uses the first reachable path and skips unreachable network drives quickly.

Example file layout in the GitHub repository root:

```text
index.html
.nojekyll
latest.json
family-browser/bootstrap-index.json
family-browser/bootstrap.json
KKY_Tool_Revit_v2.03.exe
```
