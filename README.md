# Planogram Reset App

Mobile-first planogram reset tracker for GitHub Pages.

## Publish with GitHub Pages

1. Upload `index.html`, `app.js`, and `styles.css` to the **root** of your `Planogram` repository.
2. In GitHub, open **Settings → Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**.
4. Select **main** and **/ (root)**, then Save.
5. After GitHub finishes deploying, open the Pages site in Safari.

## Workflow

1. Convert the paper planogram to the standardized Excel workbook.
2. Open the hosted Planogram Reset App.
3. Tap **Upload Excel** and select the workbook.
4. Work by C1, C2, C3, etc., or search/filter.
5. Mark each position **Complete**, **Not Complete**, or **Missing**.
6. Progress is stored locally in that browser/device.

The importer looks for these sheets in order:
- `Reset Database`
- `Extracted Positions`
- `Import Template`
- otherwise the first usable worksheet

Required Excel fields:
- Card Name
- Column

Recommended:
- Denomination
- Category
- Row
- Position
- Confidence
- Notes
