# Claim Tracker

App web (PWA) untuk rekod belanja, simpan resit dan jana borang claim. Boleh dipasang di skrin utama iPhone dan Android.

- `index.html`, `config.js`, `manifest.webmanifest`, `sw.js`, `icons/`: paparan app (dihos di GitHub Pages).
- `apps-script/`: kod backend (Google Apps Script). Ditampal ke projek Apps Script, bukan dijalankan dari GitHub.

Data (rekod, resit, borang) disimpan dalam Google Sheet dan Google Drive pemilik, bukan dalam repo ini.
Logo borang syarikat (Aset.html) hanya disimpan dalam projek Apps Script dan dihantar kepada pengguna berdaftar selepas log masuk.
Hanya emel dalam tab "Pengguna" boleh log masuk.
