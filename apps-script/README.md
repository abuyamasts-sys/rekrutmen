## Google Sheets Storage (Apps Script)

Tujuan: simpan data CBT → Pauli (Kraepelin) → DISC ke Google Sheets (Google Drive).

### 1) Buat Spreadsheet
- Buat Google Sheet baru di Google Drive.
- Copy Spreadsheet ID dari URL (bagian antara `/d/` dan `/edit`).

### 2) Buat Apps Script & Deploy Web App
- Buka Apps Script (`script.google.com`) atau dari Google Sheet: Extensions → Apps Script.
- Paste isi `apps-script/Code.gs`.
- Isi `SPREADSHEET_ID` dan (opsional) `SHARED_SECRET`.
- Deploy → New deployment → Web app:
  - Execute as: Me
  - Who has access: Anyone (atau sesuai kebutuhan)
- Copy Web app URL.

### 3) Set konfigurasi di webapp
Edit `sheet-config.js`:
- `window.AIRTIS_SHEETS_ENDPOINT = "<WEB_APP_URL>"`
- (opsional) `window.AIRTIS_SHEETS_SECRET = "<SHARED_SECRET>"`

### Catatan
- Client menyimpan antrian di `localStorage.AIRTIS_SHEET_QUEUE` dan akan retry otomatis saat ada koneksi.
- Semua event disimpan sebagai satu baris per `kind` (`cbt`, `pauli`, `disc`) di sheet `assessments`.

