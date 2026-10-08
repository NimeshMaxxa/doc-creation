# FFS document creation

Tools for making Fire Fighting Solutions (Pvt) Ltd invoices in the company layout.

## Invoice form (website)

Published with GitHub Pages at **https://nimeshmaxxa.github.io/doc-creation/** and locked with a
username and password.

- `site/index.html` is only a login page. The real form, including the signature and stamp
  images, is stored inside it encrypted (AES-256-GCM, key from PBKDF2-SHA256), so nothing
  useful can be read from this public repository without the login.
- The plain form source and the signature/stamp images are kept privately, outside this
  repository. To publish a change, rebuild the locked page from them:

      SITE_USER=... SITE_PASSWORD=... python tools/lock_site.py <private-source-folder> site

  Rebuilding with a new password changes the login; everyone signs in again.
- Saved invoices stay in the browser you use. Use **Back up** now and then and keep the file
  (for example in Google Drive); **Restore** loads it on another computer or browser.

Every push to `main` that changes `site/` republishes the page (see `.github/workflows/pages.yml`).

## Excel + Python version

`python/` makes the same PDF from an Excel sheet. See `python/README.md`.

## Changing company details

Address, phone numbers and bank details are in the `COMPANY` block near the top of the
form source (and in `python/make_invoice.py`). The logo is `site/logo.png`.
