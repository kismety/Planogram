# Planno — Attachment + JSON Backup

## Attach Planogram
Two explicit options:
- **Take Photo** — opens the phone camera.
- **Attach File / Photo** — choose an existing image from Photos or Files.

The attached image is compressed for storage, previews in Planno, and is saved with the named planogram.

## Where the data is normally saved
Planno is hosted on GitHub Pages. Its normal saved data lives in Safari/browser **localStorage on that device**.

This includes:
- saved stores and addresses
- saved planograms
- imported Excel/PDF rows
- Complete / Missing / Not Complete statuses
- active progress

## Export ALL Data
Use **Export ALL Data (.json)** to create one backup file containing:
- all Planno localStorage records
- active planogram data
- store information
- checklist statuses
- attached planogram image

Use **Import ALL Data (.json)** to restore the backup on the same or another device/browser.
