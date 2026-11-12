# Auto-Refresh Excel SharePoint

Script Python ringan untuk **me-refresh file Excel yang tersimpan di SharePoint** secara otomatis via Windows Task Scheduler. Script ini membuka file menggunakan protocol handler `ms-excel://`, memanggil `RefreshAll` lewat COM (win32com), lalu menutup Excel — tanpa perlu interaksi manual.

---

## Cara Kerja

```
Windows Task Scheduler
        │
        ▼
  run_refresh.bat          ← entry point scheduler
        │
        ▼
   refresh.py              ← logika utama (Python + win32com)
        │
        ├─ Buka file SharePoint via ms-excel:// protocol handler
        ├─ Tunggu Excel membuka file (OPEN_WAIT_SEC detik)
        ├─ Attach ke Excel yang sedang berjalan (GetActiveObject)
        ├─ Panggil wb.RefreshAll() di semua workbook yang terbuka
        ├─ Tutup workbook & quit Excel
        └─ Tulis log ke worker.log (append)
```

Setiap kali dijalankan, semua event ditulis ke `worker.log` (append), sehingga riwayat lengkap refresh tersimpan dan mudah diaudit.

---

## Prasyarat

| Kebutuhan | Keterangan |
|---|---|
| **Windows** | Script hanya berjalan di Windows (COM / win32com) |
| **Microsoft Excel** | Harus terinstall dan sudah login ke akun Microsoft 365 |
| **Python 3.8+** | Disarankan 3.10 ke atas |
| **pywin32** | `pip install pywin32` |
| **python-dotenv** | `pip install python-dotenv` |

Install semua dependensi sekaligus:

```bash
pip install pywin32 python-dotenv
```

---

## Konfigurasi

### 1. Ganti Target Link (Wajib)

Buka [`refresh.py`](refresh.py) dan ubah nilai `TARGET_LINK` di bagian atas file:

```python
# refresh.py  –– baris ~16
TARGET_LINK = (
    "https://o365indosat-my.sharepoint.com/:x:/g/personal/..."
    "/NAMA-FILE?e=TOKEN"
)
```

> **Cara dapat link:** Buka file Excel di SharePoint → klik **Share** → **Copy link** → pilih *Anyone with the link* atau *People in your organization*.

### 2. Timeout Buka File (Opsional)

Default: **60 detik**. Jika koneksi lambat atau file besar, naikkan nilainya lewat environment variable (atau file `.env`):

```env
# .env
EXCEL_OPEN_TIMEOUT_SEC=90
```

Atau langsung set sebelum menjalankan:

```bat
set EXCEL_OPEN_TIMEOUT_SEC=90
python refresh.py
```

---

## Menjalankan Manual

```bat
run_refresh.bat
```

atau langsung:

```bat
python refresh.py
```

Script akan mencetak log ke konsol **dan** mengappend ke `worker.log` di folder yang sama.

---

## Setup Windows Task Scheduler

Ikuti langkah berikut untuk menjadwalkan refresh otomatis:

### Langkah 1 – Buka Task Scheduler

Tekan `Win + R` → ketik `taskschd.msc` → Enter.

### Langkah 2 – Buat Task Baru

1. Klik **Action → Create Task** (bukan *Create Basic Task*).
2. Tab **General**:
   - *Name*: `Excel Auto Refresh`
   - Centang **Run whether user is logged on or not** (jika ingin jalan di background).
   - Centang **Run with highest privileges**.
3. Tab **Triggers** → klik **New**:
   - *Begin the task*: `On a schedule`
   - Pilih frekuensi (contoh: **Daily**, ulangi setiap **1 jam** selama **1 hari**).
   - Atur waktu mulai sesuai kebutuhan.
4. Tab **Actions** → klik **New**:
   - *Action*: `Start a program`
   - *Program/script*: isi path lengkap ke `run_refresh.bat`, contoh:
     ```
     C:\Users\fresn\OneDrive\Documents\INSP-Project\Auto-Refresh\run_refresh.bat
     ```
   - *Start in*: isi folder yang sama (agar `worker.log` dibuat di tempat yang benar):
     ```
     C:\Users\fresn\OneDrive\Documents\INSP-Project\Auto-Refresh
     ```
5. Tab **Conditions**: hapus centang *Start only if the computer is on AC power* jika pakai laptop.
6. Klik **OK** → masukkan password Windows jika diminta.

### Contoh Frekuensi Umum

| Kebutuhan | Pengaturan Trigger |
|---|---|
| Setiap jam (jam kerja) | Daily, repeat every **1 hour** for **8 hours**, start 08:00 |
| Setiap 30 menit | Daily, repeat every **30 minutes** for **1 day** |
| Sekali sehari pagi | Daily, 07:00 |
| Setiap hari kerja | Weekly, Senin–Jumat |

---

## Monitoring via `worker.log`

File `worker.log` dibuat otomatis di folder project dan **diappend** setiap kali script berjalan. Gunakan untuk memastikan refresh berjalan normal atau menemukan error.

Contoh isi log yang normal:

```
2026-11-12 08:00:01,123 [INFO] [MainThread] === Worker jalan: 2026-11-12T08:00:01 ===
2026-11-12 08:00:01,124 [INFO] [MainThread] Target file: https://o365indosat-my.sharepoint.com/...
2026-11-12 08:00:01,125 [INFO] [MainThread] Menunggu 60s agar file kebuka...
2026-11-12 08:01:01,200 [INFO] [MainThread] Workbook terbuka di Excel aktif: 1
2026-11-12 08:01:01,201 [INFO] [MainThread] RefreshAll -> NamaFile.xlsx
2026-11-12 08:01:06,500 [INFO] [MainThread] Menutup workbook: NamaFile.xlsx
2026-11-12 08:01:07,100 [INFO] [MainThread] Excel ditutup.
2026-11-12 08:01:07,101 [INFO] [MainThread] === Done ===
```

### Tanda-tanda Error di Log

| Pesan di Log | Penyebab | Solusi |
|---|---|---|
| `Tidak bisa attach ke Excel yang jalan` | Excel tidak terbuka / file gagal dibuka | Pastikan sudah login Microsoft 365, cek `TARGET_LINK` |
| `Tidak ada workbook terbuka` | Link expired atau token SharePoint kadaluarsa | Perbarui `TARGET_LINK` |
| `RefreshAll error` | Data source tidak bisa diakses | Cek koneksi internet / kredensial data source |

---

## Struktur File

```
Auto-Refresh/
├── refresh.py          # Script utama (logika refresh)
├── run_refresh.bat     # Entry point untuk Task Scheduler
├── worker.log          # Log append (dibuat otomatis saat pertama jalan)
├── .env                # (opsional) konfigurasi environment variable
└── README.md           # Dokumentasi ini
```

---

## Troubleshooting

**Excel tidak terbuka sama sekali**
- Pastikan `TARGET_LINK` adalah link sharing yang valid dan belum expired.
- Coba buka link di browser — pastikan langsung membuka Excel (bukan halaman login).

**Script jalan tapi log tidak dibuat**
- Pastikan kolom *Start in* di Task Scheduler sudah diisi dengan path folder project, bukan dikosongkan.

**Muncul dialog login Microsoft di Excel**
- Login manual satu kali ke akun yang sama dengan link SharePoint, lalu token akan tersimpan.

**Task Scheduler tidak jalan saat user logout**
- Di tab *General*, centang **Run whether user is logged on or not** dan pastikan password disimpan.

