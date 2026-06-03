# app.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
from datetime import datetime

from SearchEngine import SearchEngine


def format_path(file_path, root_dir):

    relative_path = os.path.relpath(
        file_path,
        root_dir
    )

    return " → ".join(
        relative_path.split(os.sep)
    )


class ExcelApp:

    def __init__(self, root):

        self.root = root

        self.root.title("Data Retrieval & Scrubbing Tool")

        self.root.geometry("980x600")

        self.root.configure(bg="#FFFFFF")

        self.folder_path = None

        self.output_folder = None

        self.entries = []

        self.columns = []

        self.search_engine = None

        self.final_df = None
        
        self.last_checks = []
        self.last_summary_results = []

        self.build_ui()

    # ================= UI =================

    def build_ui(self):

        style = ttk.Style()

        style.theme_use("clam")

        style.configure(
            "TFrame",
            background="#FFFFFF"
        )

        style.configure(
            "TLabel",
            background="#FFFFFF",
            foreground="#1F1F1F"
        )
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor="#FFFFFF",   # background bar
            background="#5B3CC4",    # filled color
            # bordercolor="#5B3CC4",
            lightcolor="#5B3CC4",
            darkcolor="#5B3CC4"
        )

        # ================= HEADER =================

        header = ttk.Frame(self.root)

        header.pack(
            fill="x",
            padx=12,
            pady=10
        )

        ttk.Label(
            header,
            text="DRS",
            font=("Segoe UI", 18, "bold"),
            foreground="#5B3CC4"
        ).pack()

        # ================= MAIN =================

        main = ttk.Frame(self.root)

        main.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        # ================= LEFT =================

        left = ttk.Frame(main)

        left.pack(
            side="left",
            fill="y",
            padx=(0, 10)
        )

        ttk.Label(
            left,
            text="Source",
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w")

        self.source_label = ttk.Label(
            left,
            text="No folder selected"
        )

        self.source_label.pack(
            fill="x",
            pady=3
        )

        tk.Button(
            left,
            text="Select Source Folder",
            command=self.select_folder,
            bg="#5B3CC4",
            fg="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            padx=10,
            pady=6
        ).pack(
            fill="x",
            pady=5
        )

        ttk.Label(
            left,
            text="Output",
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w")

        self.output_label = ttk.Label(
            left,
            text="No output selected"
        )

        self.output_label.pack(
            fill="x",
            pady=3
        )

        tk.Button(
            left,
            text="Select Output Folder",
            command=self.select_output_folder,
            bg="#5B3CC4",
            fg="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            padx=10,
            pady=6
        ).pack(
            fill="x",
            pady=5
        )

        # ================= FILE NAME =================

        self.file_name_entry = tk.Entry(left)

        self.file_name_entry.pack(
            fill="x",
            pady=8
        )

        placeholder = "Enter excel name with .xlsx"

        self.file_name_entry.insert(
            0,
            placeholder
        )

        self.file_name_entry.config(
            fg="grey"
        )

        def on_focus_in(e):

            if self.file_name_entry.get() == placeholder:

                self.file_name_entry.delete(
                    0,
                    tk.END
                )

                self.file_name_entry.config(
                    fg="black"
                )

        def on_focus_out(e):

            if not self.file_name_entry.get():

                self.file_name_entry.insert(
                    0,
                    placeholder
                )

                self.file_name_entry.config(
                    fg="grey"
                )

        self.file_name_entry.bind(
            "<FocusIn>",
            on_focus_in
        )

        self.file_name_entry.bind(
            "<FocusOut>",
            on_focus_out
        )

        # ================= PROGRESS =================

        self.progress = ttk.Progressbar(
            left,
            mode="determinate",
            style="Custom.Horizontal.TProgressbar"
        )

        self.progress.pack(
            fill="x",
            pady=10
        )

        # ================= CENTER =================

        center = ttk.Frame(main)

        center.pack(
            side="left",
            fill="both",
            expand=True,
            padx=10
        )

        ttk.Label(
            center,
            text="No of conditions to verify",
            font=("Segoe UI", 11, "bold")
        ).pack(anchor="w")

        self.num_entry = ttk.Entry(center)

        self.num_entry.pack(
            fill="x",
            pady=5,
            # padx=5
        )

        ttk.Button(
            center,
            text="Create Fields",
            command=self.create_fields
        ).pack(
            fill="x",
            pady=5
        )

        container = ttk.Frame(center)

        container.pack(
            fill="both",
            expand=True
        )

        canvas = tk.Canvas(
            container,
            bg="#FFFFFF",
            highlightthickness=0
        )

        scrollbar = ttk.Scrollbar(
            container,
            orient="vertical",
            command=canvas.yview
        )

        self.field_frame = ttk.Frame(canvas)

        self.canvas_window = canvas.create_window(
            (0, 0),
            window=self.field_frame,
            anchor="nw"
        )

        # IMPORTANT: proper resize handling
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event):
            canvas.itemconfig(self.canvas_window, width=event.width +7)  # 👈 GAP FIX

        canvas.bind("<Configure>", on_canvas_configure)
        self.field_frame.bind("<Configure>", on_frame_configure)

        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=(0, 10))  # 👈 GAP FIX
        scrollbar.pack(side="right", fill="y", padx=(8, 0))

        # ================= RIGHT =================

        right = ttk.Frame(main)

        right.pack(
            side="right",
            fill="both",
            expand=True
        )

        ttk.Label(
            right,
            text="Results",
            font=("Segoe UI", 11, "bold")
        ).pack(anchor="w")

        result_container = ttk.Frame(right)

        result_container.pack(
            fill="both",
            expand=True
        )

        self.output = scrolledtext.ScrolledText(
            result_container,
            font=("Consolas", 10),
            bg="#FFFFFF",
            fg="#1F1F1F",
            wrap="none"
        )

        scrollbar_y = ttk.Scrollbar(
            result_container,
            orient="vertical",
            command=self.output.yview
        )

        self.output.configure(
            yscrollcommand=scrollbar_y.set
        )

        self.output.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar_y.pack(
            side="right",
            fill="y"
        )

        # ================= BUTTONS =================

        btn_frame = ttk.Frame(right)

        btn_frame.pack(
            fill="x",
            pady=10
        )

        self.run_btn = tk.Button(
            btn_frame,
            text="RUN SEARCH",
            command=self.start_search,
            bg="#5B3CC4",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            pady=8
        )

        self.run_btn.pack(
            side="left",
            expand=True,
            fill="x",
            padx=5
        )

        self.export_btn = tk.Button(
            btn_frame,
            text="EXPORT EXCEL",
            command=self.export_excel,
            bg="#5B3CC4",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            pady=8
        )

        self.export_btn.pack(
            side="left",
            expand=True,
            fill="x",
            padx=5
        )

    # ================= SOURCE =================

    def select_folder(self):

        path = filedialog.askdirectory()

        if not path:
            return

        self.folder_path = path

        self.source_label.config(
            text=path
        )

        self.output.delete(
            "1.0",
            tk.END
        )

        self.output.insert(
            tk.END,
            "Loading columns...\n"
        )

        self.search_engine = SearchEngine(path)

        self.columns = (
            self.search_engine.get_all_columns()
        )

        self.output.insert(
            tk.END,
            f"Loaded {len(self.columns)} columns\n"
        )

    # ================= OUTPUT =================

    def select_output_folder(self):

        path = filedialog.askdirectory()

        if not path:
            return

        self.output_folder = path

        self.output_label.config(
            text=path
        )

    # ================= FIELDS =================

    def create_fields(self):

        for w in self.field_frame.winfo_children():

            w.destroy()

        self.entries.clear()

        try:

            n = int(
                self.num_entry.get()
            )

        except ValueError:

            messagebox.showerror(
                "Error",
                "Enter valid number"
            )

            return

        if not self.columns:

            messagebox.showerror(
                "Error",
                "Select source folder first"
            )

            return

        for _ in range(n):

            row = ttk.Frame(self.field_frame)

            row.pack(
                fill="x",
                pady=5,padx=5
            )

            # COLUMN

            combo = ttk.Combobox(
                row,
                values=self.columns,
                width=25
            )

            combo.pack(
                side="left",
                expand=True,
                fill="x",padx=2
            )

            # OPERATOR

            operator_combo = ttk.Combobox(row,
                    values=["equals", "not equals", "greater than", "less than", "in"],
                    width=8, state="readonly"
                    )

            operator_combo.set("equals")

            operator_combo.pack(
                side="left",
                padx=5
            )

            # VALUE

            entry = ttk.Entry(row)

            entry.pack(
                side="left",
                expand=True,
                fill="x",
                padx=2
            )

            self.entries.append(
                (
                    combo,
                    operator_combo,
                    entry
                )
            )

    # ================= SEARCH =================

    def start_search(self):

        threading.Thread(
            target=self.run_search,
            daemon=True
        ).start()

    def run_search(self):

        self.output.delete(
            "1.0",
            tk.END
        )

        checks = []
        self.last_checks = checks.copy()

        for c, op, v in self.entries:

            col = c.get().strip()

            operator = op.get().strip()

            val = v.get().strip()

            if col and val:

                checks.append({
                    "column": col,
                    "operator": operator,
                    "value": val
                })

        if not checks:

            messagebox.showerror(
                "Error",
                "No conditions entered"
            )

            return

        self.output.insert(
            tk.END,
            "Scanning...\n\n"
        )

        def update_progress(cur, total):

            value = int(
                (cur / total) * 100
            )

            self.root.after(
                0,
                lambda: self.progress.configure(
                    value=value
                )
            )

        try:

            summary_results, self.final_df = (
                self.search_engine.run_search(
                    checks,
                    update_progress
                )
            )
            self.last_summary_results = summary_results

        except Exception as e:

            messagebox.showerror(
                "Search Error",
                str(e)
            )

            return

        if not summary_results:

            self.output.insert(
                tk.END,
                "No matches found.\n"
            )

            return

        for file, result in summary_results:

            self.output.insert(
                tk.END,
                f"{format_path(file, self.folder_path)} → {result}\n"
            )

        self.progress["value"] = 100
        
    def write_export_log(self, status):
        if not self.output_folder:
            return None
        logs_dir = os.path.join(self.output_folder, "logs")
        os.makedirs(logs_dir, exist_ok=True)

        timestamp = datetime.now().strftime("%d%m%Y%H%M%S")
        log_file_name = f"SearchSummary_{timestamp}.txt"
        log_path = os.path.join(logs_dir, log_file_name)

        try:
            with open(log_path, "w", encoding="utf-8") as log:
                log.write("=" * 80 + "\n")
                log.write(f"Export Time : {datetime.now():%Y-%m-%d %H:%M:%S}\n")
                log.write(f"Status      : {status}\n")
                log.write(f"Source Path : {self.folder_path}\n")
                log.write(f"Output File : {self.final_df.shape if self.final_df is not None else 'None'}\n")

            return log_path

        except Exception as e:
            print(f"Log write failed: {e}")
            return None
        

    # ================= EXPORT =================

    def export_excel(self):

        if self.final_df is None or self.final_df.empty:

            messagebox.showerror(
                "Error",
                "No data to export"
            )

            return

        if not self.output_folder:

            messagebox.showerror(
                "Error",
                "Select output folder"
            )

            return

        file_name = (
            self.file_name_entry.get().strip()
        )

        if (
            file_name ==
            "Enter excel name with .xlsx"
            or file_name == ""
        ):

            messagebox.showerror(
                "Error",
                "Enter valid file name"
            )

            return

        if not file_name.endswith(".xlsx"):

            file_name += ".xlsx"

        output_path = os.path.join(
            self.output_folder,
            file_name
        )

        try:
            # export excel
            self.search_engine.export_results(self.final_df, output_path)

            # create log + get file path
            log_path = self.write_export_log("SUCCESS")

            msg = f"Excel saved:\n{output_path}"

            # show log path if created
            if log_path:
                msg += f"\n\nLog saved:\n{log_path}"

            messagebox.showinfo("Success", msg)

        except Exception as e:
            log_path = self.write_export_log(f"FAILED - {e}")

            msg = f"Export failed:\n{e}"

            if log_path:
                msg += f"\n\nLog saved:\n{log_path}"

            messagebox.showerror("Export Error", msg)


# ================= RUN =================

if __name__ == "__main__":

    root = tk.Tk()

    app = ExcelApp(root)

    root.mainloop()
