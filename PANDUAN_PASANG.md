# Panduan pasang Claim Tracker versi GitHub Pages

Data (Sheet, resit, borang) kekal sama. Link app lama terus boleh diguna sehingga versi baru siap.
Akaun GitHub: **frapmuse**. Alamat app nanti: **https://frapmuse.github.io/claim-tracker/**

---

## Langkah 1: Cipta repo GitHub (2 minit)

1. Buka https://github.com/new
2. Repository name: `claim-tracker`
3. Pilih **Public**. GitHub Pages percuma hanya untuk repo public. Repo ini cuma ada kod app. Ia tiada resit, data, emel atau logo syarikat.
4. Jangan tanda apa-apa kotak lain. Tekan **Create repository**.
5. Bagitahu Claude "repo dah siap". Claude akan masukkan semua kod.

## Langkah 2: Cipta Google Client ID (untuk butang "Sign in with Google") (5 minit)

1. Buka https://console.cloud.google.com dan log masuk dengan akaun Google boss.
2. Di atas sekali, tekan pemilih projek, kemudian **New Project**. Namakan `Claim Tracker` dan tekan **Create**. Pastikan projek ini yang dipilih.
3. Di kotak carian atas, taip **Google Auth Platform** dan buka. Tekan **Get started**:
   - App name: `Claim Tracker`, User support email: emel boss, kemudian **Next**
   - Audience: **External**, kemudian **Next**
   - Contact information: emel boss, kemudian **Next**, tanda setuju, dan tekan **Create**
4. Menu kiri, pergi **Audience**. Di bahagian **Test users**, tekan **Add users**, masukkan emel boss dan emel Amalina, kemudian **Save**.
5. Menu kiri, pergi **Clients** > **Create client**:
   - Application type: **Web application**
   - Name: `Claim Tracker`
   - **Authorized JavaScript origins** > Add URI: `https://frapmuse.github.io`
   - Tekan **Create**
6. Salin **Client ID** (yang berakhir dengan `.apps.googleusercontent.com`) dan hantar kepada Claude.

## Langkah 3: Kemas kini Apps Script (5 minit)

Buka projek Apps Script yang sama (script.google.com > Claim Tracker).

1. **Code.gs**: tekan Cmd+A, Delete, tampal isi `apps-script/Code.gs` yang baru, kemudian Save.
2. **appsscript.json**: tekan Cmd+A, Delete, tampal isi `apps-script/appsscript.json` yang baru, kemudian Save.
3. **Aset.html** dan **Index.html**: biar seperti biasa, jangan ubah.
4. Ikon gear (**Project Settings**) di kiri, kemudian scroll ke **Script Properties** > **Add script property**:
   - `GOOGLE_CLIENT_ID` = Client ID dari Langkah 2
   - `APP_URL` = `https://frapmuse.github.io/claim-tracker/`
   - Tekan **Save script properties**
5. Kembali ke editor. Pilih fungsi `setup` dan tekan **Run**. Benarkan (Allow) akses baru bila diminta. Setup selamat dijalankan semula; ia tidak padam apa-apa.
6. **Deploy > New deployment** (BUKAN Manage deployments):
   - Ikon gear > **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Tekan **Deploy**, kemudian salin **Web app URL** (berakhir dengan `/exec`) dan hantar kepada Claude.

"Anyone" diperlukan supaya app di GitHub boleh bercakap dengan skrip ini. Skrip tetap menolak sesiapa yang tiada dalam tab Pengguna. Setiap permintaan mesti ada sesi log masuk Google yang sah.

## Langkah 4: Claude siapkan app

Claude masukkan Web app URL dan Client ID ke dalam `config.js`, kemudian hidupkan GitHub Pages.

## Langkah 5: Pasang di telefon

- **iPhone**: buka https://frapmuse.github.io/claim-tracker/ dalam **Safari**, kemudian tekan **Sign in with Google**. Lepas tu tekan butang **Kongsi** > **Add to Home Screen**. Ikon kucing akan keluar sendiri.
- **Android**: buka link yang sama dalam **Chrome**, kemudian **Sign in with Google**. Lepas tu tekan menu **⋮** > **Install app** (atau **Add to Home screen**).
- Log masuk kekal 60 hari di telefon tersebut. Butang **Log keluar** ada dalam Profil.
- Amalina pun buat cara sama, guna akaun Google dia sendiri.

Lepas semua ok, shortcut lama dan link lama boleh dibuang.
