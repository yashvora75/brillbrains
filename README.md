# BrillBrains Website

Static website for BrillBrains Consultants Pvt. Ltd.

## Folder Structure

- `index.html` is the home page.
- Page folders such as `about`, `services`, `contact`, and service-specific folders contain clean URL versions of each page.
- Root `.html` files such as `about.html` are lightweight redirects for legacy links.
- `assets/css/styles.css` is the stylesheet manifest.
- `assets/css/01-foundation.css` through `assets/css/06-effects.css` hold the ordered CSS layers.
- `assets/js/main.js` contains navigation, reveal, counter, and hero motion behavior.
- `assets/images/brand` contains logo and favicon assets.
- `assets/images/pages` contains compressed WebP page hero images.
- `sitemap.xml` and `robots.txt` are included for search engine discovery.

## Editing Notes

- Keep public page folders and root redirect files in sync when URL plans change.
- Add new images inside the closest matching `assets/images` subfolder.
- Keep the CSS import order in `assets/css/styles.css`; later files intentionally override earlier layers.
- Update shared navigation and footer markup consistently across all public pages.
