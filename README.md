# Frame Studio for Word v1.1

Frame Studio is a Word task-pane add-in for inserting designed text frames and ASCII art. The v1.1 update adds three work areas:

- **Presets:** Full page frame, header frame, bottom page frame, and ASCII full page frame.
- **Insert Page:** Simple, mid, and complex ASCII shape libraries with live previews.
- **Draw Page:** Choose ASCII complexity, corner shape, line style, title, and body to draw a custom frame.

## Files to upload to GitHub Pages

Upload these files to the root of the `word-frame-studio` repository and keep the `assets` folder:

```text
manifest.xml
taskpane.html
taskpane.css
taskpane.js
commands.html
assets/icon-16.png
assets/icon-32.png
assets/icon-80.png
```

The manifest is already configured for:

```text
https://ahm-mar.github.io/word-frame-studio/
```

## Updating the GitHub site

1. Open the `ahm-mar/word-frame-studio` repository on GitHub.
2. Upload or replace the files listed above. Do not rename the `assets` folder.
3. Commit the changes to the branch used by GitHub Pages, normally `main`.
4. Wait for GitHub Pages to finish publishing.
5. Open the add-in again in Word. If the old pane remains, close Word completely and reopen it. If necessary, remove and reinstall the add-in manifest.

## Windows installer

`install-word-frame-studio.bat` is a current-user installer patterned after the original installer. Keep it in the same folder as `manifest.xml`, then double-click it. It copies the manifest into `%APPDATA%\\FrameStudio` and registers the add-in under the current user's Office developer sideload key.

If your organization blocks developer sideloading, use the Microsoft 365 admin center instead. An administrator can deploy this manifest URL through **Settings → Integrated apps → Add-ins → Deploy Add-in**:

```text
https://ahm-mar.github.io/word-frame-studio/manifest.xml
```

## Local development

```bash
npm install
npm run install-cert
npm start
```

The local development server uses `https://localhost:3000` only for development; the committed manifest uses the GitHub Pages URL.
