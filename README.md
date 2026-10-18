# Voice Presentation Controller

A browser-based presentation viewer by **Suryansh Sharma**. Open supported PDF and PowerPoint (`.pptx`) files and navigate with voice commands or keyboard shortcuts.

## Live demo

https://offx-suryansh.github.io/Voice-Presentation-Controller/

## Developer links

- Portfolio: https://offx-suryansh.github.io/
- GitHub: https://github.com/Offx-suryansh
- LinkedIn: https://www.linkedin.com/in/offxsuryansh/
- Source repository: https://github.com/Offx-suryansh/Voice-Presentation-Controller

## Run locally

This project uses browser JavaScript modules, so run it through a local web server rather than opening `index.html` directly. For example, with Node.js installed:

```bash
npx serve .
```

Then open the local URL shown by the server.

## Deploy with GitHub Pages

1. Open repository **Settings → Pages**.
2. Choose **Deploy from a branch**.
3. Select `main` and `/(root)`, then save.
4. Keep `index.html`, `css/`, `js/`, and `vendor/` at the repository root.

## Privacy notes

Presentation rendering and history are handled in the browser. Browser speech recognition may be processed by the browser vendor, depending on the browser implementation. Avoid using sensitive files if you are unsure about your browser's speech-recognition privacy practices.

## License

See [LICENSE](LICENSE).


## GitHub Pages deployment

This project is served as static files. The PPTX renderer is included under `vendor/pptx-browser/src/`, and PDF.js is included under `vendor/pdfjs/`, so browser imports do not depend on a `node_modules` directory being deployed. In the repository settings, choose **Deploy from a branch**, branch `main`, folder `/ (root)`.
