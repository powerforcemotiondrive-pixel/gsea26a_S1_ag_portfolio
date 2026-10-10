# GSEA English Learning Portfolio

A formative assessment portfolio application designed with a graphite and claret color scheme, tactile mesh background texture, and elevated drop shadow boxes.

## Project Structure

```
├── src/
│   └── input.css                    # Tailwind input stylesheet with @source directives
├── output.css                       # Compiled & minified Tailwind CSS output
├── gsea26a_aas_portfolio.html       # Main portfolio application
├── index.html                       # Entry-point launcher redirecting to the main application
├── imsmanifest.xml                  # IMS Content Package specification manifest for Moodle
├── package.json                     # NPM configuration & build scripts
└── moodle_portfolio_package.zip     # Ready-to-upload Moodle IMS Content Package
```

## Tailwind CLI Usage

The project uses `@tailwindcss/cli` to compile `src/input.css` into `output.css`.

- **Production Build (Minified)**:
  ```bash
  npm run build
  ```
- **Development Build (Unminified)**:
  ```bash
  npm run build:dev
  ```
- **Watch Mode (Recompiles on save)**:
  ```bash
  npm run watch
  # or
  npm run dev
  ```

## Moodle Deployment (IMS Content Package)

1. Build the latest CSS and package archive:
   ```bash
   npm run build
   npm run package:moodle
   ```
   This generates `moodle_portfolio_package.zip` with `imsmanifest.xml` positioned at the archive root.

2. In your Moodle course:
   - Turn **Edit mode** ON.
   - Click **Add an activity or resource** -> select **IMS Content Package**.
   - Set the title (e.g., *GSEA English Learning Portfolio*).
   - In the **Package file** area, upload `moodle_portfolio_package.zip`.
   - Click **Save and display**.

