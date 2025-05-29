from tkinterdnd2 import DND_FILES, TkinterDnD
from ttkbootstrap import Style
from ttkbootstrap.constants import X
from ttkbootstrap import ttk
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import os
from utils.file_handling import is_valid_excel_file
from config.sheet_definitions import REQUIRED_SHEETS

ICON_PATH = os.path.join("assets", "excel_icon.png")
# ... [unchanged imports from previous version] ...


class V8MergerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.style = Style("cosmo")  # Use a Bootstrap theme like 'flatly'
        self.title("V8 Merger")
        self.geometry("600x580")
        self.configure(bg="white")
        self.resizable(False, False)

        self.selected_files = []
        self.file_tiles = []

        self.create_widgets()

    def create_widgets(self):
        # Instruction Label
        self.instruction_label = ttk.Label(
            self,
            text="Drag and drop V8 Excel inspection files below or click Browse",
            font=("Segoe UI", 11, "bold"),
            bootstyle="dark",
        )
        self.instruction_label.pack(pady=(20, 2))

        self.sub_instruction_label = ttk.Label(
            self,
            text="Only V8-format fire inspection files are supported.",
            font=("Segoe UI", 9, "italic"),
            bootstyle="secondary",
        )
        self.sub_instruction_label.pack(pady=(0, 10))

        # Drop Zone
        self.drop_frame = ttk.Frame(
            self, style="light.TFrame", padding=10, width=500, height=260
        )
        self.drop_frame.pack_propagate(False)
        self.drop_frame.pack(padx=20, pady=(0, 10), fill=X)

        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self.handle_drop)

        self.drop_label = ttk.Label(
            self.drop_frame,
            text="Drop Excel files here",
            font=("Segoe UI", 10, "italic"),
            bootstyle="secondary",
        )
        self.drop_label.pack(pady=(0, 6))

        # File Tiles Container (inside drop zone)
        self.tiles_container = ttk.Frame(self.drop_frame)
        self.tiles_container.pack()

        # Browse Button
        self.browse_button = ttk.Button(
            self,
            text="Browse",
            command=self.browse_files,
            bootstyle="primary-outline",
            width=25,
        )
        self.browse_button.pack(pady=(10, 5))

        # Status label
        self.status_label = ttk.Label(
            self, text="", font=("Segoe UI", 9), bootstyle="secondary"
        )
        self.status_label.pack(pady=(5, 10))

        # Action Buttons Frame
        self.actions_frame = ttk.Frame(self)
        self.actions_frame.pack(pady=(10, 20))

        self.clear_button = ttk.Button(
            self.actions_frame,
            text="Clear All",
            command=self.clear_all_files,
            bootstyle="secondary-outline",
            width=20,
        )
        self.clear_button.grid(row=0, column=0, padx=10)

        self.merge_button = ttk.Button(
            self.actions_frame,
            text="▶ Start Merge",
            command=self.start_merge,
            bootstyle="success",
            width=25,
        )
        self.merge_button.grid(row=0, column=1, padx=10)
        self.merge_button.grid_remove()

        self.progress_var = tk.DoubleVar()
        self.progress_label = ttk.Label(self, text="Ready")

        self.progress_bar = ttk.Progressbar(
            self,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
            bootstyle="primary",
        )

        # Load Excel icon
        if os.path.exists(ICON_PATH):
            img = Image.open(ICON_PATH).resize((32, 32))
            self.icon_image = ImageTk.PhotoImage(img)
        else:
            self.icon_image = None

    def update_progress(self, percentage, message):
        self.progress_var.set(percentage)
        self.progress_label.config(text=message)
        self.update_idletasks()  # Force immediate update

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
                    bootstyle="danger",
                )
                continue
            elif cleaned.endswith(".xlsx"):
                if is_valid_excel_file(cleaned, REQUIRED_SHEETS):
                    self.selected_files.append(cleaned)
                    self.add_file_tile(cleaned)
                    added_count += 1
                else:
                    self.status_label.config(
                        text=(
                            f"File Error: Missing required "
                            f"sheets in {os.path.basename(cleaned)}"
                        ),
                        bootstyle="danger",
                    )
            else:
                self.status_label.config(
                    text=(
                        f"File Error: Not a valid "
                        f"Excel file → {os.path.basename(cleaned)}"
                    ),
                    bootstyle="danger",
                )

        if added_count > 0:
            self.status_label.config(
                text=f"{len(self.selected_files)} files selected", bootstyle="secondary"
            )
            self.merge_button.grid()

    def add_file_tile(self, filepath):
        frame = ttk.Frame(self.tiles_container, padding=6, style="light.TFrame")
        frame.pack(pady=4, fill=X)

        if self.icon_image:
            icon = ttk.Label(frame, image=self.icon_image)
            icon.image = self.icon_image
            icon.pack(side="left", padx=(0, 10))

        label = ttk.Label(frame, text=os.path.basename(filepath), font=("Segoe UI", 9))
        label.pack(side="left")

        remove_btn = ttk.Button(
            frame,
            text="✕",
            command=lambda: self.remove_file(filepath, frame),
            bootstyle="danger-outline round",
            width=3,
        )
        remove_btn.pack(side="right", padx=5)

        self.file_tiles.append(frame)

    def remove_file(self, filepath, frame):
        if filepath in self.selected_files:
            self.selected_files.remove(filepath)
            frame.destroy()
            self.status_label.config(
                text=f"{len(self.selected_files)} files selected", bootstyle="secondary"
            )
        if not self.selected_files:
            self.merge_button.grid_remove()

    def clear_all_files(self):
        if messagebox.askyesno(
            "Confirm", "Are you sure you want to remove all selected files?"
        ):
            for frame in self.file_tiles:
                frame.destroy()
            self.selected_files.clear()
            self.file_tiles.clear()
            self.status_label.config(text="No files selected", bootstyle="secondary")
            self.merge_button.grid_remove()

    def start_merge(self):
        from core.merger import merge
        import shutil

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
            # Show the progress bar and label
            self.progress_label.pack(side="bottom", fill="x", padx=20, pady=5)
            self.progress_bar.pack(side="bottom", fill="x", padx=10)
            self.progress_var.set(0)
            self.progress_label.config(text="Opening template...")
            self.update_idletasks()
            conflicts = merge(
                self.selected_files, save_path, progress_callback=self.update_progress
            )
            self.progress_var.set(100)
            # self.progress_bar.configure(bootstyle="success")
            self.progress_label.config(text="✅ Merge complete!")
            self.update_idletasks()
            unique_conflicts = sorted(set(conflicts))
            conflict_summary = (
                "\n".join(unique_conflicts)
                if unique_conflicts
                else "No conflicts found!"
            )

            messagebox.showinfo(
                "Merge Complete",
                f"✅ Merge complete!\n\nConflict Summary:\n{conflict_summary}",
            )

            # Auto-close after 1 second (1000ms)
            self.after(50, self.destroy)
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save merged file:\n{e}")


def launch_app():
    app = V8MergerApp()
    app.mainloop()
