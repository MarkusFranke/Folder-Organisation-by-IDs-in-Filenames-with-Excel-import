import os
import shutil
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook

class FileOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel ID File Organizer")

        self.source_folder = ""
        self.destination_folder = ""
        self.ids = []
        self.last_dir = os.path.expanduser("~")

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Select Source Folder").grid(row=0, column=0, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.browse_source_folder).grid(row=0, column=1, padx=10, pady=10)
        self.source_folder_label = tk.Label(self.root, text="No folder selected")
        self.source_folder_label.grid(row=0, column=2, padx=10, pady=10)

        tk.Label(self.root, text="Enter IDs (comma separated or paste from Excel)").grid(row=1, column=0, padx=10, pady=10)
        self.ids_entry = tk.Entry(self.root, width=50)
        self.ids_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(self.root, text="Or Load IDs from Excel File").grid(row=2, column=0, padx=10, pady=10)
        tk.Button(self.root, text="Load Excel File", command=self.load_excel_file).grid(row=2, column=1, padx=10, pady=10)

        tk.Label(self.root, text="Select Destination Folder").grid(row=3, column=0, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.browse_destination_folder).grid(row=3, column=1, padx=10, pady=10)
        self.destination_folder_label = tk.Label(self.root, text="No folder selected")
        self.destination_folder_label.grid(row=3, column=2, padx=10, pady=10)

        tk.Button(self.root, text="Organize Files", command=self.organize_files).grid(row=4, columnspan=3, padx=10, pady=20)

    def browse_source_folder(self):
        self.source_folder = filedialog.askdirectory(initialdir=self.last_dir).replace("\\", "/")
        if self.source_folder:
            self.last_dir = os.path.dirname(self.source_folder)
            self.source_folder_label.config(text=self.source_folder)
        print(f"Selected source folder: {self.source_folder}")

    def browse_destination_folder(self):
        self.destination_folder = filedialog.askdirectory(initialdir=self.last_dir).replace("\\", "/")
        if self.destination_folder:
            self.last_dir = os.path.dirname(self.source_folder)
            self.destination_folder_label.config(text=self.destination_folder)
        print(f"Selected destination folder: {self.destination_folder}")

    def load_excel_file(self):
        excel_file_path = filedialog.askopenfilename(initialdir=self.last_dir, filetypes=[("Excel files", "*.xlsx;*.xls")]).replace("\\", "/")
        if excel_file_path:
            self.last_dir = os.path.dirname(excel_file_path)
            
            try:
                df = pd.read_excel(excel_file_path, dtype=str, header=None)
                # Assume the IDs are in the first column
                self.ids = df.iloc[:, 0].dropna().tolist()
                self.ids_entry.delete(0, tk.END)
                self.ids_entry.insert(0, ", ".join(self.ids))
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel file: {e}")

    def organize_files(self):
        if not self.ids:
            self.ids = [id.strip() for id in self.ids_entry.get().split(",")]
        if not self.source_folder or not self.destination_folder or not self.ids:
            messagebox.showerror("Error", "Please ensure all fields are filled out correctly.")
            return

        self.organize_files_by_ids()
        messagebox.showinfo("Success", "Files have been organized successfully!")

    def organize_files_by_ids(self):
        if not os.path.exists(self.destination_folder):
            os.makedirs(self.destination_folder)

        matched_files = set()
        matched_ids = []
        unmatched_ids = []
        unmatched_files = []
        matched_filenames = []

        id_patterns = [re.compile(rf'^[^.\d]*{id}[^.\d]*\.png$', re.IGNORECASE) for id in self.ids]
        id_patterns_nef = [re.compile(rf'^[^.\d]*{id}[^.\d]*\.nef$', re.IGNORECASE) for id in self.ids]

        for root_dir, _, files in os.walk(self.source_folder):
            for filename in files:
                source_path = os.path.join(root_dir, filename).replace("\\", "/")
                matched = False
                for id, pattern in zip(self.ids, id_patterns):
                    if pattern.match(filename):
                        destination_path = os.path.join(self.destination_folder, filename).replace("\\", "/")
                        try:
                            shutil.copy2(source_path, destination_path)
                            matched_files.add(source_path)
                            matched_filenames.append(filename)
                            matched_ids.append(id)
                            matched = True
                            break
                        except PermissionError:
                            messagebox.showerror("Error", f"Permission denied: Cannot copy file {source_path}")
                            return
                if not matched:
                    unmatched_files.append(filename)
        for root_dir, _, files in os.walk(self.source_folder):
            for filename in files:
                source_path = os.path.join(root_dir, filename).replace("\\", "/")
                matched = False
                for id, pattern_nef in zip(self.ids, id_patterns_nef):
                    if pattern_nef.match(filename) and id not in matched_ids:
                        destination_path = os.path.join(self.destination_folder, filename).replace("\\", "/")
                        try:
                            shutil.copy2(source_path, destination_path)
                            matched_files.add(source_path)
                            matched_filenames.append(filename)
                            matched_ids.append(id)
                            matched = True
                            break
                        except PermissionError:
                            messagebox.showerror("Error", f"Permission denied: Cannot copy file {source_path}")
                            return
                if not matched:
                    unmatched_files.append(filename)
                    
        unmatched_files = set(unmatched_files)            
        unmatched_files = [file for file in unmatched_files if file not in matched_filenames]
        
        # Find unmatched IDs
        for id in self.ids:
            if id not in matched_ids:
                unmatched_ids.append(id)

        self.save_report_to_excel(matched_filenames, unmatched_ids, unmatched_files)

    def save_report_to_excel(self, matched_filenames, unmatched_ids, unmatched_files):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Report"

        sheet.append(["Matched Files", "Unmatched ProduktNr", "Unmatched Files"])
        max_len = max(len(matched_filenames), len(unmatched_ids), len(unmatched_files))
        for i in range(max_len):
            matched_filename = matched_filenames[i] if i < len(matched_filenames) else ""
            unmatched_id = unmatched_ids[i] if i < len(unmatched_ids) else ""
            unmatched_file = unmatched_files[i] if i < len(unmatched_files) else ""
            sheet.append([matched_filename, unmatched_id, unmatched_file])

        excel_path = os.path.join(self.destination_folder, "!report.xlsx").replace("\\", "/")
        workbook.save(excel_path)
        print(f"Report saved to {excel_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileOrganizerApp(root)
    root.mainloop()
