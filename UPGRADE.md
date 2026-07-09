# LTS Portal — upgrade to v5

## 1. Backend
Paste the new `Code.gs` → **Deploy → Manage deployments → edit → Version: New version**.
(Never "New deployment" — the /exec URL would change.)

## 2. Run ONE function from the editor

```
repairAll()
```

It is idempotent and does everything:
- renames `K400` → `J966` across every sheet
- creates the `LTSVersions`, `Diagrams`, `Annotations` tables
- adds `Shares.level`, `Uploads.versionId`
- seeds a v1 baseline for every existing LTS and binds attachments to it
- removes duplicate substation rows

If anything still looks wrong, run `diagnose()` and send me the log output.

## 3. Frontend
Upload `Index.html` to GitHub as **`index.html`** (lowercase). Hard-reload (Ctrl+Shift+R).
Login screen must show `build v5-2026-07-09`.
