# DataTable Editor

Browser-based data table editor for early game planning and schema design.

## Project Layout

```text
DataTable_Editor/
├─ index.html
├─ README.md
└─ assets/
   ├─ css/
   │  └─ styles.css
   └─ js/
      └─ app.js
```

## What Lives Where

- `index.html`
  Entry point for GitHub Pages and the static app shell.
- `assets/css/styles.css`
  App-wide styling, layout, ERD visuals, and editor component styles.
- `assets/js/app.js`
  State, rendering, validation, import/export, ERD, and editor interaction logic.

## GitHub Pages

This repository is structured to work cleanly with GitHub Pages from the repository root.

- Pages source: `main` branch
- Folder: `/ (root)`
- Expected URL:
  `https://yang-yun-seok.github.io/DataTable_Editor/`

Because asset paths are relative, the app works under the repository subpath without extra build setup.

## Local Preview

Run any simple static server from the repository root.

```powershell
py -m http.server 4173
```

Then open:

```text
http://127.0.0.1:4173/
```

## Notes

- This is a static frontend project. No backend or build step is required.
- Project data is stored in browser local storage unless exported.
