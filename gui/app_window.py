from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import filedialog, messagebox
import tkinter as tk
import win32com.client
import os


class V8MergerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("V8 Merger")
        self.geometry("520x500")
        self.configure(bg="white")
        self.resizable(False, False)

        self.selected_files = []

        self.create_widgets()

    def create_widgets(self):
        # Title/Instruction
        self.instruction_label = tk.Label(
            self,
            text="Drag and drop Excel files below or click Browse",
            bg="white",
            font=("Segoe UI", 11, "bold"),
            fg="#333",
        )
        self.instruction_label.pack(pady=(20, 10))

        # Drop zone
        self.drop_frame = tk.Frame(self, bg="#e9ecef", bd=2, relief="ridge", height=100)
        self.drop_frame.pack(padx=20, fill=tk.X)
        self.drop_frame.pack_propagate(False)

        self.drop_label = tk.Label(
            self.drop_frame,
            text="Drop Excel files here",
            bg="#e9ecef",
            font=("Segoe UI", 10, "italic"),
            fg="#495057",
        )
        self.drop_label.pack(expand=True)

        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self.handle_drop)

        # Browse button
        self.browse_button = tk.Button(
            self,
            text="Browse",
            command=self.browse_files,
            font=("Segoe UI", 10),
            width=25,
            bg="#2c7be5",
            fg="white",
            relief="raised",
            bd=2,
        )
        self.browse_button.pack(pady=(15, 5))

        # Merge button
        self.merge_button = tk.Button(
            self,
            text="Start Merge",
            command=self.start_merge,
            font=("Segoe UI", 10),
            width=25,
            bg="#38c172",
            fg="white",
            relief="raised",
            bd=2,
        )
        self.merge_button.pack(pady=(5, 15))

        # File list
        self.file_listbox = tk.Listbox(
            self, height=8, width=60, font=("Segoe UI", 9), bd=1, relief="solid"
        )
        self.file_listbox.pack(padx=20, pady=(0, 20))

        # Status label
        self.status_label = tk.Label(
            self, text="", bg="white", font=("Segoe UI", 9), fg="gray"
        )
        self.status_label.pack()

    def browse_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        self.add_files(files)

    def handle_drop(self, event):
        files = self.tk.splitlist(event.data)
        self.add_files(files)

    def resolve_shortcut(self, path):
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(path)
        return shortcut.TargetPath

    def add_files(self, files):
        print(f"Adding files: {files}")  # Add this line

        for f in files:
            print(f"Checking: {f}")  # Debug

            # Resolve shortcuts
            if f.endswith(".lnk"):
                f = self.resolve_shortcut(f)

            cleaned = f.strip('"')
            if cleaned.endswith(".xlsx") and cleaned not in self.selected_files:
                print(f"Accepted: {f}")  # Debug
                self.selected_files.append(cleaned)
                self.file_listbox.insert(tk.END, os.path.basename(cleaned))

        self.status_label.config(text=f"{len(self.selected_files)} files selected")

    def start_merge(self):
        if not self.selected_files:
            messagebox.showwarning("No files", "Please select Excel files to merge.")
            return
        messagebox.showinfo("Merge", f"Merging {len(self.selected_files)} files...")
        self.status_label.config(text="Merge complete!")


def launch_app():
    app = V8MergerApp()
    app.mainloop()
