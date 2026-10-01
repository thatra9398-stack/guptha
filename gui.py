import os
import sys
import threading
import tempfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from openpyxl import Workbook

from ais_to_excel import (
    read_ais,
    AisParserError,
    sanitize_sheet_name,
    sheet_base_name,
    write_summary_sheet,
    write_category_sheet,
)

ctk.set_appearance_mode("System")  # Follow Windows Dark/Light mode
ctk.set_default_color_theme("blue")


class AISConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AIS to Excel Converter")
        self.geometry("540x580")
        self.resizable(False, False)

        self.selected_json_path = None

        self._build_ui()

    def _build_ui(self):
        # Header Container
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(padx=24, pady=(24, 12), fill="x")

        title_lbl = ctk.CTkLabel(
            header_frame,
            text="AIS to Excel Converter",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        )
        title_lbl.pack(fill="x")

        sub_lbl = ctk.CTkLabel(
            header_frame,
            text="Convert Income Tax AIS downloads into clean multi-sheet Excel workbooks.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
            anchor="w",
        )
        sub_lbl.pack(fill="x", pady=(2, 0))

        # Main Card Frame
        card = ctk.CTkFrame(self, corner_radius=12)
        card.pack(padx=24, pady=8, fill="both", expand=True)

        # File Selection Section
        file_lbl = ctk.CTkLabel(
            card,
            text="1. Select AIS JSON File",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        )
        file_lbl.pack(padx=20, pady=(16, 6), fill="x")

        file_select_frame = ctk.CTkFrame(card, fg_color="transparent")
        file_select_frame.pack(padx=20, pady=0, fill="x")

        self.file_path_entry = ctk.CTkEntry(
            file_select_frame,
            placeholder_text="No file selected...",
            state="readonly",
        )
        self.file_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        browse_btn = ctk.CTkButton(
            file_select_frame,
            text="Browse",
            width=80,
            command=self._browse_file,
        )
        browse_btn.pack(side="right")

        # Credentials Section
        cred_lbl = ctk.CTkLabel(
            card,
            text="2. Enter Credentials",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        )
        cred_lbl.pack(padx=20, pady=(18, 6), fill="x")

        # PAN Input
        pan_lbl = ctk.CTkLabel(card, text="PAN (10 characters):", font=ctk.CTkFont(size=12), anchor="w")
        pan_lbl.pack(padx=20, pady=(2, 0), fill="x")
        self.pan_entry = ctk.CTkEntry(card, placeholder_text="ABCDE1234F")
        self.pan_entry.pack(padx=20, pady=(0, 8), fill="x")

        # DOB Input
        dob_lbl = ctk.CTkLabel(card, text="Date of Birth (DDMMYYYY):", font=ctk.CTkFont(size=12), anchor="w")
        dob_lbl.pack(padx=20, pady=(2, 0), fill="x")
        self.dob_entry = ctk.CTkEntry(card, placeholder_text="01011990", show="•")
        self.dob_entry.pack(padx=20, pady=(0, 12), fill="x")

        # Status & Progress
        self.status_label = ctk.CTkLabel(
            card,
            text="🔒 All processing runs 100% locally on your PC.",
            font=ctk.CTkFont(size=11),
            text_color="gray",
        )
        self.status_label.pack(padx=20, pady=(4, 8))

        self.progress = ctk.CTkProgressBar(card)
        self.progress.pack(padx=20, pady=(0, 16), fill="x")
        self.progress.set(0)

        # Convert Action Button
        self.convert_btn = ctk.CTkButton(
            self,
            text="Convert to Excel",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=42,
            corner_radius=8,
            command=self._start_conversion,
        )
        self.convert_btn.pack(padx=24, pady=(8, 24), fill="x")

    def _browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select AIS JSON File",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")],
        )
        if file_path:
            self.selected_json_path = Path(file_path)
            self.file_path_entry.configure(state="normal")
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, str(self.selected_json_path))
            self.file_path_entry.configure(state="readonly")

    def _start_conversion(self):
        json_path_str = self.file_path_entry.get().strip()
        pan = self.pan_entry.get().strip().upper()
        dob = self.dob_entry.get().strip()

        if not json_path_str or not self.selected_json_path or not self.selected_json_path.exists():
            messagebox.showerror("Error", "Please select a valid AIS JSON file.")
            return

        if not pan or len(pan) != 10 or not pan.isalnum():
            messagebox.showerror("Error", "Please enter a valid 10-character PAN.")
            return

        if not dob or len(dob) != 8 or not dob.isdigit():
            messagebox.showerror("Error", "Please enter Date of Birth as 8 digits (DDMMYYYY).")
            return

        output_path = filedialog.asksaveasfilename(
            title="Save Output Excel File",
            defaultextension=".xlsx",
            initialfile=self.selected_json_path.stem + ".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )

        if not output_path:
            return

        self.convert_btn.configure(state="disabled")
        self.status_label.configure(text="Processing AIS JSON...", text_color="#1f538d")
        self.progress.configure(mode="indeterminate")
        self.progress.start()

        threading.Thread(
            target=self._run_conversion_thread,
            args=(self.selected_json_path, f"{pan}{dob}", Path(output_path)),
            daemon=True,
        ).start()

    def _run_conversion_thread(self, json_path: Path, password: str, output_path: Path):
        try:
            statement = read_ais(path=json_path, password=password)
            frames = statement.to_df()

            used_names = {"Summary"}
            sheet_name_by_key = {
                key: sanitize_sheet_name(sheet_base_name(key, df), used_names) if len(df) > 0 else None
                for key, df in sorted(frames.items())
            }

            wb = Workbook()
            write_summary_sheet(wb, statement, frames, sheet_name_by_key)

            for key, df in sorted(frames.items()):
                if len(df) == 0:
                    continue
                rows = df.to_dict(orient="records")
                write_category_sheet(wb, sheet_name_by_key[key], list(df.columns), rows)

            wb.save(output_path)

            self.after(
                0,
                self._conversion_success,
                f"Successfully converted {len(frames)} categories!\nSaved to: {output_path.name}",
            )
        except AisParserError as exc:
            self.after(0, self._conversion_error, f"Decryption failed: {exc}\nCheck PAN and Date of Birth.")
        except Exception as exc:
            self.after(0, self._conversion_error, f"Error during conversion: {exc}")

    def _conversion_success(self, msg: str):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(1.0)
        self.status_label.configure(text="✅ Conversion complete!", text_color="green")
        self.convert_btn.configure(state="normal")
        messagebox.showinfo("Success", msg)

    def _conversion_error(self, msg: str):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(0)
        self.status_label.configure(text="❌ Conversion failed.", text_color="red")
        self.convert_btn.configure(state="normal")
        messagebox.showerror("Conversion Error", msg)


if __name__ == "__main__":
    app = AISConverterApp()
    app.mainloop()
