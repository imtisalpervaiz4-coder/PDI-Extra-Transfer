"""
PDI Extra Transfer Assistant
Excel columns:
  A = ITEM NO
  E = CODE
  G = XFER QTY

PDI sequence for every row:
  Paste ITEM NO -> TAB x2 -> paste Excel CODE -> TAB
  -> paste XFER QTY -> TAB x3

Hotkeys:
  F9  = Start / Resume
  F10 = Pause
  F11 = Stop
"""

import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

from openpyxl import load_workbook
import pyautogui
import pyperclip

try:
    import keyboard
    HOTKEYS_AVAILABLE = True
except ImportError:
    keyboard = None
    HOTKEYS_AVAILABLE = False

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05

START_DELAY = 8
AFTER_ITEM_DELAY = 1.50
AFTER_PASTE_DELAY = 0.55
TAB_DELAY = 0.40
ROW_DELAY = 0.90


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
            raise ValueError("Expected ITEM NO header in Excel column A.")
        if str(ws["E1"].value or "").strip().upper() != "CODE":
            raise ValueError("Expected CODE header in Excel column E.")
        if str(ws["G1"].value or "").strip().upper() != "XFER QTY":
            raise ValueError("Expected XFER QTY header in Excel column G.")

        rows = []
        for r in range(2, ws.max_row + 1):
            item = as_text(ws[f"A{r}"].value)
            if not item:
                continue

            code = as_text(ws[f"E{r}"].value)
            qty = as_text(ws[f"G{r}"].value)

            if not code:
                raise ValueError(f"Code is blank in Excel row {r} (column E).")
            if not qty:
                raise ValueError(f"XFER QTY is blank in Excel row {r} (column G).")

            rows.append({
                "excel_row": r,
                "item": item,
                "code": code,
                "qty": qty,
            })

        if not rows:
            raise ValueError("No item rows found.")
        return rows
    finally:
        wb.close()


class TransferDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("PDI Extra Transfer Assistant")
        self.root.geometry("700x520")
        self.root.minsize(620, 470)

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

        ttk.Label(outer, text="PDI EXTRA TRANSFER",
                  font=("Segoe UI", 19, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="Excel-to-PDI automation • ITEM NO + Excel CODE + XFER QTY",
            font=("Segoe UI", 9)
        ).pack(anchor="w", pady=(2, 15))

        file_frame = ttk.LabelFrame(outer, text="Excel workbook", padding=12)
        file_frame.pack(fill="x")
        ttk.Entry(file_frame, textvariable=self.excel_path).pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Button(file_frame, text="Browse…", command=self.choose_file).pack(side="right")
        ttk.Button(file_frame, text="Load Excel", command=self.load_excel).pack(
            side="right", padx=(0, 8)
        )

        info = ttk.LabelFrame(outer, text="Excel mapping and PDI sequence", padding=12)
        info.pack(fill="x", pady=12)
        ttk.Label(
            info,
            text=(
                "Excel A = ITEM NO     |     Excel E = CODE     |     Excel G = XFER QTY\n"
                "PDI: Paste Item → TAB ×2 → Paste Excel Code → TAB → Paste Qty → TAB ×3"
            ),
            justify="left"
        ).pack(anchor="w")

        status_frame = ttk.LabelFrame(outer, text="Progress", padding=12)
        status_frame.pack(fill="x")
        ttk.Label(status_frame, textvariable=self.status,
                  font=("Segoe UI", 11, "bold")).pack(anchor="w")
        ttk.Label(status_frame, textvariable=self.detail).pack(anchor="w", pady=(4, 8))
        self.progress = ttk.Progressbar(status_frame, mode="determinate", maximum=100)
        self.progress.pack(fill="x")
        ttk.Label(status_frame, textvariable=self.progress_text).pack(
            anchor="e", pady=(5, 0)
        )

        controls = ttk.Frame(outer)
        controls.pack(fill="x", pady=(15, 8))
        ttk.Button(controls, text="▶  START / RESUME  (F9)",
                   command=self.start_resume).pack(
            side="left", fill="x", expand=True, padx=(0, 5)
        )
        ttk.Button(controls, text="Ⅱ  PAUSE  (F10)",
                   command=self.pause).pack(
            side="left", fill="x", expand=True, padx=5
        )
        ttk.Button(controls, text="■  STOP  (F11)",
                   command=self.stop).pack(
            side="left", fill="x", expand=True, padx=(5, 0)
        )

        hotkey_text = "Hotkeys: F9 Start/Resume   •   F10 Pause   •   F11 Stop"
        if not HOTKEYS_AVAILABLE:
            hotkey_text += "\nHotkeys unavailable: install the keyboard package."
        ttk.Label(outer, text=hotkey_text, foreground="#555").pack(anchor="w", pady=5)
        ttk.Label(
            outer,
            text=("Before starting: open PDI and click the FIRST EMPTY Number cell. "
                  "You have 8 seconds to switch to PDI."),
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
        if self.running:
            messagebox.showwarning("Busy", "Stop the current transfer before loading another file.")
            return
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showwarning("Select Excel", "Choose your Excel workbook first.")
            return
        try:
            self.rows = load_rows(path)
        except Exception as exc:
            messagebox.showerror("Excel error", str(exc))
            return

        self.current_index = 0
        self.progress["value"] = 0
        self.progress_text.set(f"0 / {len(self.rows)}")
        self.status.set("Excel loaded successfully.")
        self.detail.set(f"{len(self.rows)} item rows ready • Code is read from column E.")
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
        self.detail.set(f"Paused at item {self.current_index + 1}.")

    def stop(self):
        if not self.running:
            self.status.set("Stopped.")
            return
        self.stop_event.set()
        self.pause_event.set()
        self.status.set("Stopping…")
        self.detail.set("Stopping after the current keyboard action.")

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
        self.progress["value"] = completed / total * 100 if total else 0
        self.progress_text.set(f"{completed} / {total}")
        self.detail.set(
            f"Excel row {row['excel_row']} • Item {row['item']} • "
            f"Code {row['code']} • Qty {row['qty']}"
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

                # Exact PDI sequence:
                # ITEM NO -> TAB -> TAB -> Excel column E CODE -> TAB
                # -> XFER QTY -> TAB -> TAB -> TAB
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
                self.ui(self.detail.set,
                        f"Stopped at {self.current_index} of {total} completed rows.")
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
