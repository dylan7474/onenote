# OneNote Web

A browser-based notebook with automatic server-side persistence. Notes, imports,
drawings, notebooks, sections, and pages are saved to a JSON data file and loaded
again whenever the application is opened.

## Run locally

```bash
PORT=3020 DATA_FILE=./data/state.json node server.js
```

Open <http://localhost:3020>. Do not open `index.html` directly: the application
uses the server's `/api/state` endpoint to save and restore data.

## Deploy with Docker

```bash
./deploy.sh 3020
```

The deployment stores application state in the `onenote-data` Docker volume, so
rebuilding or replacing the application container does not delete saved notes.
