# BrillBrains Website

Static website for BrillBrains Consultants Pvt. Ltd.

## Folder Structure

- `index.html` and the other root `.html` files are the public pages.
- `assets/css/styles.css` is the stylesheet manifest.
- `assets/css/01-foundation.css` through `assets/css/06-effects.css` hold the ordered CSS layers.
- `assets/js/main.js` contains navigation, reveal, counter, and hero motion behavior.
- `assets/images/brand` contains logo and favicon assets.
- `assets/images/pages` contains page-specific photos.
- `assets/images/illustrations` contains illustration assets.
- `assets/images/team` contains team/founder artwork.
- `archive` contains older standalone pages kept for reference.

## Editing Notes

- Keep public page files at the root unless the URL plan changes.
- Add new images inside the closest matching `assets/images` subfolder.
- Keep the CSS import order in `assets/css/styles.css`; later files intentionally override earlier layers.
- Update shared navigation and footer markup consistently across all public pages.
