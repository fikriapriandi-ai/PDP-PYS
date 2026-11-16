# Policy Deployment Plan (PDP) Pys

Dashboard web satu file (`index.html`) untuk PDP 2H 2026: Ringkasan, Matrix (X-matrix L1/L2/L3), Monitor KPI, Objective, Prioritas, Penjabaran, PIC, dan Riwayat. Tanpa build step, tinggal dibuka di browser.

## Fitur
- Matrix Hoshin Kanri Level 1–3 yang bisa diklik, dengan filter PIC/objective, zoom, dan layar penuh
- Monitor KPI dengan filter, sort, lampu status per bulan, dan tren
- Tema Otomatis/Terang/Gelap, ekspor CSV/PDF, impor Excel, edit data langsung, log perubahan

## Deploy (GitHub Pages)
1. Push semua file ini ke branch `main`.
2. Settings → Pages → Source: **GitHub Actions**. Workflow `.github/workflows/pages.yml` akan men-deploy otomatis.

## Memperbarui data dari Excel
```bash
pip install openpyxl
python scripts/build_data.py path/ke/New_PDP_2H2026.xlsx   # menulis scripts/data.json
python scripts/inject_data.py                              # menanam data ke index.html
```
Atau cukup pakai tombol **Impor Excel** di dashboard (disimpan di browser pengguna itu saja).

## Catatan
- `index.html` memuat seluruh data PDP. Gunakan repo **private** bila data bersifat rahasia (Pages untuk repo private butuh paket berbayar).
- Impor Excel dan ekspor PDF memuat SheetJS/jsPDF dari cdnjs, jadi perlu internet.
- File `.xlsx` di-ignore agar tidak ikut ter-commit.
