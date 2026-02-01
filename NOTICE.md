# Notices

This repository is MIT licensed, but it interacts with and/or bundles components from other projects.

## Bundled files

- `extension/httpd.sys.mjs`
  - Source: Mozilla (Thunderbird)
  - License: MPL-2.0
  - This file is copied (as-is) to provide a minimal embedded HTTP server.

## Runtime dependencies (not bundled)

- TbSync (`tbsync@jobisoft.de`) by John Bieling
  - Repo: https://github.com/jobisoft/TbSync
  - License: MPL-2.0

- Provider for Exchange ActiveSync (`eas4tbsync@jobisoft.de`) by John Bieling
  - Repo: https://github.com/jobisoft/EAS-4-TbSync
  - License: MPL-2.0

- Thunderbird Calendar APIs (Mozilla)
  - License: MPL-2.0

This project does not redistribute TbSync or EAS-4-TbSync; it only checks for their presence and calls their public/internal APIs inside Thunderbird.
