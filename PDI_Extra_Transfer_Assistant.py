"""
PDI Extra Transfer Assistant
Excel columns used:
  A = ITEM NO
  E = PACKAGE CODE (read separately for each row)
  G = XFER QTY
PDI sequence for EVERY Excel data row:
  Paste ITEM NO -> TAB -> TAB -> paste Excel column E code -> TAB -> paste XFER QTY
  -> TAB -> TAB -> TAB (skip Cost, skip Retail, next row)

Hotkeys:
  F9  = Start / Resume
  F10 = Pause
  F11 = Stop

Keep this file and your Excel workbook wherever you like. Select the workbook
using the dashboard's Browse button. Review the PDI transaction before saving.
"""

import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

try:
    from openpyxl import load_workbook
    import pyautogui
    import pyperclip
except ImportError as exc:
    raise SystemExit(
        f"Missing package: {exc}\n\n"
        "Install dependencies with:\n"
        "py -m pip install openpyxl pyautogui pyperclip keyboard"
    )

try:
    import keyboard
    HOTKEYS_AVAILABLE = True
except ImportError:
    keyboard = None
    HOTKEYS_AVAILABLE = False


# ---------- SETTINGS: preserve the exact PDI field sequence ----------
START_DELAY = 8
AFTER_ITEM_DELAY = 1.50
AFTER_PASTE_DELAY = 0.55
TAB_DELAY = 0.40
ROW_DELAY = 0.90

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05


def as_text(value):
    if value is None:
        return ""
    try:
        number = float(value)
        if number.is_integer():
            return str(int(number))
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def load_rows(excel_file):
    path = Path(excel_file)
    if not path.exists():
        raise FileNotFoundError(f"Cannot find Excel file:\n{path}")

    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb[wb.sheetnames[0]]

        if str(ws["A1"].value or "").strip().upper() != "ITEM NO":
            raise ValueError("Expected header ITEM NO in Excel column A.")
        if str(ws["G1"].value or "").strip().upper() != "XFER QTY":
            raise ValueError("Expected header XFER QTY in Excel column G.")

        rows = []
        for row_number in range(2, ws.max_row + 1):
            item = ws[f"A{row_number}"].value
            if item is None or str(item).strip() == "":
                continue
            code = ws[f"E{row_number}"].value
            qty = ws[f"G{row_number}"].value
            if code is None or str(code).strip() == "":
                raise ValueError(
                    f"Package code is blank in Excel column E, row {row_number}."
                )
            rows.append({
                "excel_row": row_number,
                "item": as_text(item),
                "code": as_text(code),
                "qty": as_text(qty),
            })

        if not rows:
            raise ValueError("No item rows found in the Excel file.")
        return rows
    finally:
        wb.close()


class TransferDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("PDI Extra Transfer Assistant")
        self.root.geometry("690x510")
        self.root.minsize(620, 460)

        self.excel_path = tk.StringVar()
        self.status = tk.StringVar(value="Choose your Excel file to begin.")
        self.detail = tk.StringVar(value="Waiting for Excel file")
        self.progress_text = tk.StringVar(value="0 / 0")
        self.rows = []
        self.worker = None
        self.pause_event = threading.Event()
        self.stop_event = threading.Event()
        self.running = False
        self.paused = False
        self.current_index = 0
        self.hotkeys_registered = False

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self._register_hotkeys()

    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer, text="PDI EXTRA TRANSFER",
            font=("Segoe UI", 19, "bold")
        ).pack(anchor="w")
        ttk.Label(
            outer,
            text="Automated Excel-to-PDI entry • ITEM NO + CODE + XFER QTY",
            font=("Segoe UI", 9)
        ).pack(anchor="w", pady=(2, 15))

        file_frame = ttk.LabelFrame(outer, text="Excel workbook", padding=12)
        file_frame.pack(fill="x")
        file_entry = ttk.Entry(file_frame, textvariable=self.excel_path)
        file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(file_frame, text="Browse…", command=self.choose_file).pack(side="right")
        ttk.Button(file_frame, text="Load Excel", command=self.load_excel).pack(
            side="right", padx=(0, 8)
        )

        info = ttk.LabelFrame(outer, text="Transfer rules (fixed)", padding=12)
        info.pack(fill="x", pady=12)
        ttk.Label(
            info,
            text=(
                "A: ITEM NO     |     E: PACKAGE CODE     |     G: XFER QTY\n"
                "PDI: Paste Item → TAB ×2 → Paste Excel Code → TAB → Paste Qty → TAB ×3"
            ),
            justify="left"
        ).pack(anchor="w")

        status_frame = ttk.LabelFrame(outer, text="Progress", padding=12)
        status_frame.pack(fill="x")
        ttk.Label(status_frame, textvariable=self.status, font=("Segoe UI", 11, "bold")).pack(anchor="w")
        ttk.Label(status_frame, textvariable=self.detail).pack(anchor="w", pady=(4, 8))
        self.progress = ttk.Progressbar(status_frame, mode="determinate", maximum=100)
        self.progress.pack(fill="x")
        ttk.Label(status_frame, textvariable=self.progress_text).pack(anchor="e", pady=(5, 0))

        controls = ttk.Frame(outer)
        controls.pack(fill="x", pady=(15, 8))
        self.start_button = ttk.Button(
            controls, text="▶  START / RESUME  (F9)", command=self.start_resume
        )
        self.start_button.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.pause_button = ttk.Button(
            controls, text="Ⅱ  PAUSE  (F10)", command=self.pause
        )
        self.pause_button.pack(side="left", fill="x", expand=True, padx=5)
        self.stop_button = ttk.Button(
            controls, text="■  STOP  (F11)", command=self.stop
        )
        self.stop_button.pack(side="left", fill="x", expand=True, padx=(5, 0))

        hotkey_text = "Hotkeys: F9 Start/Resume   •   F10 Pause   •   F11 Stop"
        if not HOTKEYS_AVAILABLE:
            hotkey_text += "\nHotkeys unavailable until you install: keyboard"
        ttk.Label(outer, text=hotkey_text, foreground="#555").pack(anchor="w", pady=(4, 5))
        ttk.Label(
            outer,
            text="Before starting: open PDI, click the FIRST EMPTY Number cell. "
                 "You have 8 seconds to switch to PDI.",
            wraplength=640
        ).pack(anchor="w")

    def ui(self, callback, *args):
        self.root.after(0, lambda: callback(*args))

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Select transfer Excel workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if path:
            self.excel_path.set(path)
            self.load_excel()

    def load_excel(self):
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showwarning("Select Excel", "Choose your Excel workbook first.")
            return
        try:
            rows = load_rows(path)
        except Exception as exc:
            messagebox.showerror("Excel error", str(exc))
            return

        if self.running:
            messagebox.showwarning("Busy", "Stop the current transfer before loading another file.")
            return

        self.rows = rows
        self.current_index = 0
        self.progress["value"] = 0
        self.progress_text.set(f"0 / {len(rows)}")
        self.status.set("Excel loaded successfully.")
        self.detail.set(f"{len(rows)} item rows ready • Codes read from Excel column E")
        self.pause_event.clear()
        self.stop_event.clear()

    def _register_hotkeys(self):
        if not HOTKEYS_AVAILABLE:
            return
        try:
            keyboard.add_hotkey("F9", lambda: self.ui(self.start_resume))
            keyboard.add_hotkey("F10", lambda: self.ui(self.pause))
            keyboard.add_hotkey("F11", lambda: self.ui(self.stop))
            self.hotkeys_registered = True
        except Exception as exc:
            self.status.set(f"Hotkeys could not be registered: {exc}")

    def start_resume(self):
        if not self.rows:
            messagebox.showinfo("Load Excel", "Choose and load the Excel workbook first.")
            return
        if self.running and not self.paused:
            return

        self.stop_event.clear()
        self.pause_event.set()
        self.paused = False

        if self.worker and self.worker.is_alive():
            self.status.set("Resuming transfer…")
            self.detail.set(f"Continuing from item {self.current_index + 1}")
            return

        self.running = True
        self.status.set(f"Starting in {START_DELAY} seconds…")
        self.detail.set("Click the first empty Number cell in PDI now.")
        self.worker = threading.Thread(target=self._run_transfer, daemon=True)
        self.worker.start()

    def pause(self):
        if not self.running:
            return
        self.pause_event.clear()
        self.paused = True
        self.status.set("PAUSED — press F9 or Start / Resume to continue.")
        self.detail.set(f"Progress saved in this session at item {self.current_index + 1}")

    def stop(self):
        if not self.running:
            self.status.set("Stopped.")
            return
        self.stop_event.set()
        self.pause_event.set()  # release a paused worker so it can exit
        self.status.set("Stopping…")
        self.detail.set("The current keyboard/paste action will finish, then stop.")

    def _checkpoint(self):
        while not self.pause_event.is_set():
            if self.stop_event.is_set():
                return False
            time.sleep(0.08)
        return not self.stop_event.is_set()

    def _wait(self, seconds):
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            if not self._checkpoint():
                return False
            time.sleep(min(0.08, max(0, end - time.monotonic())))
        return self._checkpoint()

    def _paste(self, value):
        if not self._checkpoint():
            return False
        pyperclip.copy(str(value))
        pyautogui.hotkey("ctrl", "v")
        return self._wait(AFTER_PASTE_DELAY)

    def _tab(self):
        if not self._checkpoint():
            return False
        pyautogui.press("tab")
        return self._wait(TAB_DELAY)

    def _set_progress(self, completed, total, row):
        percent = completed / total * 100 if total else 0
        self.progress["value"] = percent
        self.progress_text.set(f"{completed} / {total}")
        self.detail.set(
            f"Excel row {row['excel_row']} • Item {row['item']} • Code {row['code']} • Qty {row['qty']}"
        )

    def _run_transfer(self):
        total = len(self.rows)
        try:
            if self.current_index == 0 and not self._wait(START_DELAY):
                return

            while self.current_index < total:
                if not self._checkpoint():
                    break

                row = self.rows[self.current_index]
                self.ui(self._set_progress, self.current_index, total, row)

                # Exact requested PDI sequence:
                # ITEM NO -> TAB -> TAB -> row's Excel E code -> TAB -> XFER QTY -> TAB x3
                if not self._paste(row["item"]):
                    break
                if not self._wait(AFTER_ITEM_DELAY):
                    break
                if not self._tab() or not self._tab():
                    break
                if not self._paste(row["code"]):
                    break
                if not self._tab():
                    break
                if not self._paste(row["qty"]):
                    break
                if not self._tab() or not self._tab() or not self._tab():
                    break
                if not self._wait(ROW_DELAY):
                    break

                self.current_index += 1
                self.ui(self._set_progress, self.current_index, total, row)

            if self.stop_event.is_set():
                self.ui(self.status.set, "STOPPED.")
                self.ui(self.detail.set, f"Stopped at {self.current_index} of {total} completed rows.")
            elif self.current_index >= total:
                self.ui(self.status.set, "COMPLETED — review PDI before saving/posting.")
                self.ui(self.detail.set, f"All {total} Excel item rows were entered.")
            else:
                self.ui(self.status.set, "PAUSED — press F9 to resume.")
        except pyautogui.FailSafeException:
            self.stop_event.set()
            self.ui(self.status.set, "EMERGENCY STOP — mouse moved to top-left corner.")
        except Exception as exc:
            self.ui(self.status.set, "ERROR — transfer stopped.")
            self.ui(self.detail.set, str(exc))
        finally:
            self.running = False
            self.paused = False
            self.pause_event.clear()

    def close(self):
        if self.running:
            self.stop_event.set()
            self.pause_event.set()
        if self.hotkeys_registered:
            try:
                keyboard.unhook_all_hotkeys()
            except Exception:
                pass
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = TransferDashboard(root)
    root.mainloop()
