# Dynamic Excel Builder & Live Auto-Fill System

A React + FastAPI application for uploading a master Excel file, creating custom working sheets, performing instant lookup-based auto-fill, editing in an Excel-like grid, preserving sessions locally, and exporting to Excel/PDF.

## Features Implemented

- Upload `.xlsx` / `.xls` master Excel files
- Select header row and re-parse
- Automatic unique column detection
- Select, rename, reorder, and remove columns
- Add, rename, reorder, and delete custom columns
- Lookup key selection from master columns
- Duplicate lookup warnings and handling modes: first, latest, popup, merge
- Excel-like grid powered by Handsontable
- Editable cells, copy/paste, multi-cell selection, row/column operations, undo/redo, frozen headers, filters/search support
- Real-time auto-fill when lookup value is entered
- Manual edit protection with configurable overwrite behavior
- Not-found row highlighting/status
- Local browser persistence using IndexedDB
- Import previously exported/generated working sheet
- Export final sheet to `.xlsx`
- Export final sheet to `.pdf`
- Docker and docker-compose support

## Project Structure

```text
dynamic-excel-builder/
  backend/
    main.py
    requirements.txt
    Dockerfile
  frontend/
    src/
      App.jsx
      api.js
      db.js
      main.jsx
      styles.css
    package.json
    Dockerfile
    nginx.conf
  docker-compose.yml
```

## Run Locally Without Docker

### Backend

```bash
cd backend
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

### Frontend

```bash
cd frontend
npm install
npm run dev
```

Open:

```text
http://localhost:5173
```

The frontend expects the backend at `http://localhost:8000` during development.

## Run With Docker

```bash
docker compose up --build
```

Open:

```text
http://localhost:3000
```

## Notes

- Session data is stored locally in the browser via IndexedDB. No external storage platform is required.
- Very large files should be parsed by the backend. Grid rendering is virtualized by Handsontable.
- Password-protected files are rejected with a friendly message.
- PDF Unicode support depends on available fonts. The backend attempts to register DejaVu Sans automatically when available.

## Latest Requested Updates

- Lookup column can now be changed directly inside **3 Work & Export**.
- Removed master columns can now be added directly inside **3 Work & Export**.
- Users can add custom columns directly inside **3 Work & Export**.
- Going back to configuration no longer deletes previously entered live-sheet data.
- Updating/regenerating the live sheet from configuration preserves existing data where columns still match.
- Lookup matching is case-insensitive.
- Partial lookup matches prompt the user before filling data from the main Excel file.
- Buttons added for clear selected cells, clear selected column, delete selected rows, refresh selected rows, and refresh all rows.
- Copy/paste limits increased for large range pasting.
- Excel/PDF exports use plain normal table styling without colored headers.

## Latest Quality Update

- External copy/paste from Excel/Google Sheets/other tables is now handled by a direct clipboard parser.
- Use Ctrl+V/Cmd+V inside the live grid, or click **Paste Clipboard**.
- Large paste ranges auto-add missing rows, and extra pasted columns become custom columns automatically.
- Live-sheet column deletion is supported with **Delete Columns** and the right-click menu.
- Live-sheet column reordering is supported by drag/drop, **Move Column** buttons, and right-click menu.
- Undo/Redo buttons were added and grid undo/redo syncs back to local session state.
- Partial lookup selection now fills the full row and also corrects the lookup field from the master Excel record.
- UI was refreshed with improved spacing, sticky toolbar, modern cards, better buttons, and clearer status messages.

## Deploy on Render

This repository includes `render.yaml`.

### Option A: Blueprint Deploy

1. Push this project to GitHub.
2. In Render, choose **New +** → **Blueprint**.
3. Select the repository.
4. Render will create:
   - `dynamic-excel-builder-api`
   - `dynamic-excel-builder-frontend`
5. If you change the backend service name, update the frontend env var `VITE_API_BASE` to your backend URL.

### Option B: Manual Deploy

Backend:

- Type: Web Service
- Environment: Docker
- Dockerfile: `backend/Dockerfile`
- Health check path: `/health`

Frontend:

- Type: Static Site
- Build command: `cd frontend && npm install && npm run build`
- Publish directory: `frontend/dist`
- Environment variable:
  - `VITE_API_BASE=https://YOUR-BACKEND-SERVICE.onrender.com`

The backend already allows CORS for the frontend.
