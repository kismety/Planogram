# Planno — Stores, Fixtures, Spinners, Reports + Google Drive

## Store-first structure
Planno now organizes work as:

**Store → Fixture / Spinner Group → Planogram Face → Reset status/report**

Save each store once with its name, address and optional Store ID, then reuse it for future resets.

## End caps and spinners
When saving a planogram, use **Fixture Structure**:
- Fixture / Spinner Group: the common physical fixture name, such as `Gift Card Spinner 1` or `End Cap 7`.
- Fixture Type: Auto, End Cap, Spinner, Main Fixture or Other.
- Face / Side: Auto, Front, Left End Cap, Right End Cap, Back, or Spinner Side A–D.

Planograms that share the same **Store + Fixture / Spinner Group** are combined on the Progress screen, while each face keeps its own Column/Row positions.

## Entire fixture report
The Progress screen includes a **Fixture / Spinner Overview** and can generate an **Entire Fixture PDF** containing:
- combined totals for all faces
- complete / missing / not-complete counts per face
- outstanding and missing items by face
- the complete item/status report for every saved face

## Attach planogram
Two explicit options remain:
- **Take Photo** — opens the phone camera.
- **Attach File / Photo** — choose an existing image from Photos or Files.

The attachment is compressed for browser storage and saved with the named planogram.

## Local saving and JSON backup
Planno is hosted on GitHub Pages. Normal auto-save uses browser **localStorage on that device**.

**Export ALL Data (.json)** saves stores, saved planograms, statuses, imported rows and attached images in one backup. **Import ALL Data (.json)** restores it.

## Optional Google Drive sync
Planno can now save a full JSON backup to Google Drive and restore the latest Drive backup. Final planogram PDFs and Entire Fixture PDFs can also be saved to:

`Planno / Stores / <Store Name> / Reports`

### One-time Google setup
Because the app runs on GitHub Pages, create a Google Cloud **OAuth 2.0 Web Client ID** with the deployed Planno site added as an authorized JavaScript origin. Paste that Client ID into the Google Drive Sync card in Planno.

The browser version uses the `drive.file` permission and a short-lived sign-in token. Local saving continues to work even when Google Drive is not connected.
