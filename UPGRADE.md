# LTS Portal — v7

## Deploy
1. Paste new `Code.gs` → **Deploy → Manage deployments → edit → Version: New version**.
2. Run `repairAll()` once (idempotent) — creates the new `Outages` and `Testing` tables
   and the `Uploads.testingId` column.
3. Upload `Index.html` to GitHub as `index.html` (lowercase). Hard-reload.
   Login screen must read `build v7-2026-07-09`.

## New in v7 — two new tabs inside every LTS

### Outages
Fields: **From**, **To** (date-time), **Duration (hh:mm, auto-calculated)**, **Findings**, **Remarks**.
Reverse-chronological list. Preview / Export (Excel, PDF, Print) at the top of the tab, Add-outage
button for editors. Entries auto-filed from a Testing entry are marked and can't be edited/deleted
directly — edit or delete the Testing entry instead, which keeps them in sync.

### Testing
Fields: **Cause**, **From**, **To**, **Duration (hh:mm, auto)**, **Findings**, **Remarks**.
Reverse-chronological list. Saving a NEW testing entry automatically files a linked **Outages**
row (same window) — a test genuinely takes the scheme off-service.
Click a testing entry to open its full detail: all fields, plus a reverse-chronological
**attachments** list with add-attachment (per-file caption prompt), each opening in the same
Drive-style lightbox viewer as LTS documents (rotate, zoom, highlight, print all work here too).

Deleting a testing entry removes its linked outage and attachments.

## Fixes carried from your last message, verified again
- Diagram/builder colours (previously "all black") — tokenised, day mode confirmed.
- Viewer/builder full-screen overlay — confirmed positioned correctly.
- `04.07.2026` style legacy dates — parsed correctly on both client and server.
