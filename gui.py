"""
GUI & Orchestration for STAF Insert Comment Tool V3.1

Notes:
- Reads data using openpyxl-only functions (safe; fast) in excel_tools.staf_logic
- Inserts Excel comments using xlwings COM in excel_tools.xlwings_comment (preserves shapes)
- Avoids duplicates and sizes legacy Notes (old-style comments) for readability.
"""
import tkinter as tk
from tkinter import filedialog, messagebox

# Logic imports (pure openpyxl read/extract work)
from excel_tools.staf_logic import (
    validate_ship_code,
    load_workbooks_readonly,
    build_comment_dict,
    extract_daily_metrics,
    detect_active_metric,
    has_surrounding_position_number,
)
# COM insertion (safe) via xlwings
from excel_tools.xlwings_comment import insert_comment_at_address

class STAFCommentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("STAF Insert Comment Tool V3.1 (xlwings)")
        self.root.geometry("640x560")

        self.source_path = None
        self.target_path = None
        self.ship_code = None

        self._build_ui()

    def _build_ui(self):
        # Ship code
        tk.Label(self.root, text="Ship Code (e.g., GR):").grid(row=0, column=0, sticky="w", padx=10, pady=(12,6))
        self.ship_code_entry = tk.Entry(self.root, width=10)
        self.ship_code_entry.grid(row=0, column=1, sticky="w", padx=10, pady=(12,6))

        # Source file
        tk.Label(self.root, text="Machine_Details.xls:").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        tk.Button(self.root, text="Browse...", command=self._pick_source).grid(row=1, column=1, sticky="w", padx=10, pady=6)

        # Target file
        tk.Label(self.root, text="STAF.xlsm:").grid(row=2, column=0, sticky="w", padx=10, pady=6)
        tk.Button(self.root, text="Browse...", command=self._pick_target).grid(row=2, column=1, sticky="w", padx=10, pady=6)

        # Preview cell (for minimal test)
        tk.Label(self.root, text="Test Cell (e.g., F12):").grid(row=3, column=0, sticky="w", padx=10, pady=6)
        self.test_cell_entry = tk.Entry(self.root, width=10)
        self.test_cell_entry.insert(0, "F12")
        self.test_cell_entry.grid(row=3, column=1, sticky="w", padx=10, pady=6)

        # Buttons
        tk.Button(self.root, text="Insert TEST comment (xlwings)",
                  command=self._insert_test_comment).grid(row=4, column=0, columnspan=2, pady=12)

        tk.Button(self.root, text="Run FULL logic (read + detect)",
                  command=self._run_full_logic).grid(row=5, column=0, columnspan=2, pady=6)

        # NEW: Batch write button
        tk.Button(self.root, text="Write comments (xlwings batch)",
                  command=self._write_comments).grid(row=6, column=0, columnspan=2, pady=6)

        # Log box (moved from row=6 to row=7)
        self.log_box = tk.Text(self.root, height=16, width=74)
        self.log_box.grid(row=7, column=0, columnspan=2, padx=10, pady=10)

        # Status (moved from row=7 to row=8)
        self.status_label = tk.Label(self.root, text="🚀 Ready.", anchor="w", relief="sunken")
        self.status_label.grid(row=8, column=0, columnspan=2, sticky="we", padx=10, pady=(0, 10))

    def _set_status(self, msg):
        self.status_label.config(text=msg)
        self.status_label.update_idletasks()

    def _log(self, msg):
        self.log_box.insert(tk.END, msg + "\\n")
        self.log_box.see(tk.END)

    def _pick_source(self):
        p = filedialog.askopenfilename(title="Select Machine_Details.xls",
                                       filetypes=[("Excel 97-2003", "*.xls"), ("Excel", "*.xlsx *.xlsm *.xls")])
        if p:
            self.source_path = p
            self._log(f"✔ Source: {p}")

    def _pick_target(self):
        p = filedialog.askopenfilename(title="Select STAF.xlsm",
                                       filetypes=[("Excel Macro-Enabled Workbook", "*.xlsm")])
        if p:
            self.target_path = p
            self._log(f"✔ Target: {p}")

    def _insert_test_comment(self):
        try:
            self._set_status("⏳ Inserting test comment via xlwings...")
            ship_code = validate_ship_code(self.ship_code_entry.get())
            if not self.target_path:
                raise ValueError("Please select a STAF.xlsm target file first.")

            cell = self.test_cell_entry.get().strip() or "F12"
            text = "STAF Tool Test\\nPosition: {}001\\nAsset Number: 61623168\\nDenom: 1¢".format(ship_code)
            summary = insert_comment_at_address(
                in_path=self.target_path,
                sheet_name="FLOOR PLAN",
                cell=cell,
                note_text=text,
                out_path=None,      # saves as *_with_Note.xlsm by default
                make_visible=False,
                autosize=True,
            )
            self._log("=== TEST INSERT SUMMARY ===")
            for k,v in summary.items():
                self._log(f"{k}: {v}")
            ok = "OK" if summary.get("shapes_intact") else "WARNING"
            self._set_status(f"✅ Test done. Shapes check: {ok}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._log(str(e))
            self._set_status("❌ Error during test insert.")

    def _run_full_logic(self):
        try:
            self._set_status("⏳ Running read/extract/detect (no write via openpyxl)...")
            ship_code = validate_ship_code(self.ship_code_entry.get())
            if not self.source_path or not self.target_path:
                raise ValueError("Please select both Machine_Details.xls and STAF.xlsm.")

            # Load workbooks read-only (we won't save with openpyxl to avoid shape loss)
            source_wb, target_wb = load_workbooks_readonly(self.source_path, self.target_path)
            self._log("✅ Files loaded (read-only).")

            # Build comments dict from Machine_Details
            comment_dict = build_comment_dict(source_wb.active, ship_code)
            self._log(f"✅ Built comment dictionary: {len(comment_dict)} entries.")

            # Extract metrics and detect which metric the floor plan displays
            machine_count = len(comment_dict)
            coin_dict, netwin_dict = extract_daily_metrics(target_wb, ship_code, machine_count, log_callback=self._log)
            floor_sheet = target_wb["FLOOR PLAN"]
            active_metric = detect_active_metric(floor_sheet, coin_dict, netwin_dict, log_callback=self._log)
            self._log(f"📌 FLOOR PLAN is displaying: {'Daily Coin In' if active_metric == 'coin_in' else 'Daily Net Win'}")

            self._set_status("✅ Logic ok. Ready to wire xlwings batch insertion next step.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._log(str(e))
            self._set_status("❌ Error while running logic.")

        self.comment_dict = comment_dict
        self.coin_dict = coin_dict
        self.netwin_dict = netwin_dict
        self.active_metric = active_metric
        self.floor_sheet = floor_sheet
        self.selected_metric = coin_dict if active_metric == 'coin_in' else netwin_dict

    def _find_placements(self, tol: float = 0.2) -> list[tuple[str, str]]:
        """
        Scan FLOOR PLAN grid and find (cell_address, comment_text) placements.
        Conditions:
          - cell value ~= metric value within tol
          - one of the 8 neighbors has the integer position number (merge-safe)
          - each pos_key inserted once only
        """
        if not hasattr(self, "selected_metric") or not hasattr(self, "comment_dict"):
            raise RuntimeError("Run FULL logic first to build dictionaries.")

        sheet = self.floor_sheet
        placements = []
        inserted_keys = set()

        # helper to get openpyxl A1 address quickly
        def addr(r, c):
            return sheet.cell(row=r, column=c).coordinate

        for r in range(1, sheet.max_row + 1):
            for c in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=r, column=c)
                val = cell.value
                if not isinstance(val, (int, float)):
                    continue

                # iterate metric dict by pos_key
                for pos_key, metric_val in self.selected_metric.items():
                    if pos_key in inserted_keys:
                        continue
                    try:
                        if abs(float(val) - float(metric_val)) < tol:
                            # verify nearby integer position (e.g., GR042 -> 42)
                            if has_surrounding_position_number(sheet, r, c, int(pos_key[-3:])):
                                comment_text = self.comment_dict.get(pos_key)
                                if comment_text:
                                    placements.append((addr(r, c), comment_text))
                                    inserted_keys.add(pos_key)
                                    break
                    except Exception:
                        continue

        self._log(f"🧭 Planned placements: {len(placements)}  (unique positions: {len(inserted_keys)})")
        return placements

    def _write_comments(self):
        try:
            self._set_status("⏳ Building placements and writing comments via xlwings...")
            if not self.target_path:
                raise ValueError("Please select a STAF.xlsm target file first.")
            if not hasattr(self, "selected_metric"):
                raise ValueError("Run FULL logic first (read + detect).")

            placements = self._find_placements(tol=0.2)

            if not placements:
                self._log("⚠ No placements found. Nothing to write.")
                self._set_status("⚠ No matches to write.")
                return

            from excel_tools.xlwings_comment import insert_comments_batch
            summary = insert_comments_batch(
                in_path=self.target_path,
                sheet_name="FLOOR PLAN",
                placements=placements,
                out_path=None,  # saves as *_with_Note.xlsm next to input
                make_visible=False,
                autosize=True,
            )

            self._log("=== BATCH WRITE SUMMARY ===")
            for k, v in summary.items():
                self._log(f"{k}: {v}")
            ok = "OK" if summary.get("shapes_intact") else "WARNING"
            self._set_status(
                f"✅ Batch write done. Shapes: {ok}. Created={summary.get('created')}, Updated={summary.get('updated')}, Skipped={summary.get('skipped')}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._log(str(e))
            self._set_status("❌ Error during batch write.")


# Allow running gui.py directly during development
if __name__ == "__main__":
    root = tk.Tk()
    app = STAFCommentApp(root)
    root.mainloop()