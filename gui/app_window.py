from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import tkinter as tk
import os
from utils.file_handling import is_valid_excel_file
from config.sheet_definitions import REQUIRED_SHEETS

ICON_PATH = os.path.join("assets", "excel_icon.png")


class V8MergerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("V8 Merger")
        self.geometry("580x600")
        self.configure(bg="white")
        self.resizable(False, False)

        self.selected_files = []
        self.file_tiles = []

        self.create_widgets()

    def create_widgets(self):
        self.instruction_label = tk.Label(
            self,
            text="Drag and drop V8 Excel inspection files below or click Browse",
            bg="white",
            font=("Segoe UI", 11, "bold"),
            fg="#333",
        )
        self.instruction_label.pack(pady=(20, 2))

        self.sub_instruction_label = tk.Label(
            self,
            text="Only V8-format fire inspection files are supported.",
            bg="white",
            font=("Segoe UI", 9, "italic"),
            fg="#6c757d",
        )
        self.sub_instruction_label.pack(pady=(0, 10))

        self.drop_frame = tk.Frame(self, bg="#f8f9fa", bd=2, relief="ridge", height=120)
        self.drop_frame.pack(padx=20, fill=tk.X)
        self.drop_frame.pack_propagate(False)

        self.drop_label = tk.Label(
            self.drop_frame,
            text="Drop Excel files here",
            bg="#f8f9fa",
            font=("Segoe UI", 10, "italic"),
            fg="#495057",
        )
        self.drop_label.pack(expand=True)

        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self.handle_drop)

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
        self.browse_button.pack(pady=(10, 10))

        self.grid_frame = tk.Frame(self, bg="white")
        self.grid_frame.pack(padx=20, pady=(10, 5))

        self.status_label = tk.Label(
            self, text="", bg="white", font=("Segoe UI", 9), fg="gray"
        )
        self.status_label.pack(pady=5)

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
        self.merge_button.pack(pady=(5, 5))
        self.merge_button.pack_forget()

        self.clear_button = tk.Button(
            self,
            text="Clear All",
            command=self.clear_all_files,
            font=("Segoe UI", 10),
            width=25,
            bg="#6c757d",
            fg="white",
            relief="raised",
            bd=2,
        )
        self.clear_button.pack(pady=(5, 15))

        if os.path.exists(ICON_PATH):
            img = Image.open(ICON_PATH).resize((32, 32))
            self.icon_image = ImageTk.PhotoImage(img)
        else:
            self.icon_image = None

    def browse_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        self.add_files(files)

    def handle_drop(self, event):
        files = self.tk.splitlist(event.data)
        self.add_files(files)

    def add_files(self, files):
        added_count = 0
        for f in files:
            cleaned = f.strip('"')
            if cleaned in self.selected_files:
                self.status_label.config(
                    text=f"File Error: Duplicate file '{os.path.basename(cleaned)}'",
                    fg="red",
                )
                continue
            elif cleaned.endswith(".xlsx"):
                if is_valid_excel_file(cleaned, REQUIRED_SHEETS):
                    self.selected_files.append(cleaned)
                    self.add_file_tile(cleaned)
                    added_count += 1
                else:
                    self.status_label.config(
                        text=f"File Error: Missing required sheets in "
                        f"{os.path.basename(cleaned)}'",
                        fg="red",
                    )
            else:
                self.status_label.config(
                    text=f"File Error: Not a valid Excel file → "
                    f"{os.path.basename(cleaned)}",
                    fg="red",
                )

        if added_count > 0:
            self.status_label.config(
                text=f"{len(self.selected_files)} files selected", fg="gray"
            )
            self.merge_button.pack()

    def add_file_tile(self, filepath):
        frame = tk.Frame(
            self.grid_frame, bg="white", bd=1, relief="solid", padx=6, pady=6
        )
        frame.pack(side="top", pady=4)

        if self.icon_image:
            icon = tk.Label(frame, image=self.icon_image, bg="white")
            icon.image = self.icon_image
            icon.pack(side="left", padx=(0, 10))

        label = tk.Label(
            frame, text=os.path.basename(filepath), bg="white", font=("Segoe UI", 9)
        )
        label.pack(side="left")

        remove_btn = tk.Button(
            frame,
            text="X",
            command=lambda: self.remove_file(filepath, frame),
            font=("Segoe UI", 8),
            bg="#dc3545",
            fg="white",
            relief="flat",
            padx=5,
        )
        remove_btn.pack(side="right", padx=5)

        self.file_tiles.append(frame)

    def remove_file(self, filepath, frame):
        if filepath in self.selected_files:
            self.selected_files.remove(filepath)
            frame.destroy()
            self.status_label.config(
                text=f"{len(self.selected_files)} files selected", fg="gray"
            )
        if not self.selected_files:
            self.merge_button.pack_forget()

    def clear_all_files(self):
        if messagebox.askyesno(
            "Confirm", "Are you sure you want to remove all selected files?"
        ):
            for frame in self.file_tiles:
                frame.destroy()
            self.selected_files.clear()
            self.file_tiles.clear()
            self.status_label.config(text="No files selected", fg="gray")
            self.merge_button.pack_forget()

    def start_merge(self):
        from core.merger import merge
        import shutil
        import os
        from tkinter import filedialog, messagebox

        TEMPLATE_FILENAME = "Annual ULC Template - CAN,ULC-S536-19 v8.xlsx"
        template_subfolder = (
            r"Cantec Fire Alarms\Cantec Office - "
            r"Documents\Cantec\Operations\Templates\Report Templates\Log Templates"
        )
        user_profile = os.environ.get("USERPROFILE")
        TEMPLATE_PATH = os.path.join(
            user_profile, template_subfolder, TEMPLATE_FILENAME
        )

        if not os.path.exists(TEMPLATE_PATH):
            messagebox.showerror(
                "Template Missing", f"Template not found:\n{TEMPLATE_PATH}"
            )
            return

        if not self.selected_files:
            messagebox.showwarning("No files", "Please select Excel files to merge.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Merged Report As",
        )

        if not save_path:
            return

        try:
            shutil.copyfile(TEMPLATE_PATH, save_path)
        except Exception as e:
            messagebox.showerror("File Error", f"Could not copy template:\n{e}")
            return

        try:
            merge(self.selected_files, save_path)

            messagebox.showinfo("Success", f"Merged file saved to:\n{save_path}")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save merged file:\n{e}")


def launch_app():
    app = V8MergerApp()
    app.mainloop()
