import os
import time
import logging
from datetime import datetime

import win32com.client as win32
import pythoncom
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# File SharePoint yang akan di-refresh
# ---------------------------------------------------------------------------

TARGET_LINK = (
    "https://o365indosat-my.sharepoint.com/:x:/g/personal/2304007262_ioh_co_id"
    "/IQCu8nnX6ZIgR7mXOz7Y0Vs2AWTfgUBX4TTOqe1-qT4nw-A?e=HbKgcr"
)

# Berapa detik tunggu setelah protocol handler dipanggil sebelum RefreshAll.
OPEN_WAIT_SEC = int(os.environ.get("EXCEL_OPEN_TIMEOUT_SEC", "60"))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s",
    handlers=[
        logging.FileHandler("worker.log"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    pythoncom.CoInitialize()
    try:
        now = datetime.now()
        log.info(f"=== Worker jalan: {now.isoformat()} ===")
        log.info(f"Target file: {TARGET_LINK}")

        # Tembak protocol handler -> Excel yang sudah jalan akan buka file SharePoint
        safe_link = TARGET_LINK.replace("|", "^|")
        os.system(f"start ms-excel:ofe^|u^|{safe_link}")

        # Tunggu flat -- Excel butuh waktu buka file dari SharePoint
        log.info(f"Menunggu {OPEN_WAIT_SEC}s agar file kebuka...")
        time.sleep(OPEN_WAIT_SEC)

        # Attach ke Excel yang sedang jalan (bukan buat instance baru)
        try:
            excel = win32.GetActiveObject("Excel.Application")
        except Exception as e:
            log.error(f"Tidak bisa attach ke Excel yang jalan: {e}")
            return

        n = excel.Workbooks.Count
        log.info(f"Workbook terbuka di Excel aktif: {n}")
        if n == 0:
            log.error("Tidak ada workbook terbuka. Pastikan login Microsoft aktif / link valid.")
            return

        workbooks = list(excel.Workbooks)
        for wb in workbooks:
            log.info(f"RefreshAll -> {wb.Name}")
            try:
                wb.RefreshAll()
            except Exception as e:
                log.warning(f"RefreshAll error ({wb.Name}): {e}")

        time.sleep(5)  # kasih waktu rendering setelah refresh

        # Tutup semua workbook dan quit Excel
        for wb in workbooks:
            log.info(f"Menutup workbook: {wb.Name}")
            try:
                wb.Close(SaveChanges=True)
            except Exception as e:
                log.warning(f"Gagal tutup workbook ({wb.Name}): {e}")

        try:
            excel.Quit()
            log.info("Excel ditutup.")
        except Exception as e:
            log.warning(f"Gagal quit Excel: {e}")

        log.info("=== Done ===")

    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()