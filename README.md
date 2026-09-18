# Kenari Barbershop — Vercel + Google Sheets

## Struktur
- `index.html` — website utama
- `api/booking.js` — API Vercel sebagai proxy ke Google Apps Script
- `Code.gs` — backend Google Apps Script yang menulis ke Google Sheets
- `.env.example` — contoh environment variable

## 1. Google Apps Script
1. Buka Google Sheet Kenari Barbershop.
2. Extensions → Apps Script.
3. Ganti kode dengan isi `Code.gs`.
4. Save.
5. Deploy → New deployment.
6. Type: Web app.
7. Execute as: Me.
8. Who has access: Anyone.
9. Deploy.
10. Salin URL yang berakhir `/exec`.

## 2. Vercel
Upload/deploy folder project ini ke Vercel.
Di Vercel → Project → Settings → Environment Variables:
- Name: `GAS_URL`
- Value: URL Web App Google Apps Script yang berakhir `/exec`

Lalu Redeploy.

## 3. Tes
Setelah deploy, buka:
`https://DOMAIN-VERCEL-ANDA.vercel.app/api/booking?action=health`

Harus mendapatkan JSON:
`{"success":true,"message":"API Kenari Barbershop aktif."}`

Setelah itu coba reservasi. Data akan masuk ke sheet `Reservasi`.

## Catatan
Browser tidak lagi mengakses Google Apps Script secara langsung.
Alurnya:
Browser → Vercel `/api/booking` → Google Apps Script → Google Sheets

Ini menghindari masalah CORS/fetch langsung yang sebelumnya menghasilkan `NetworkError`.
