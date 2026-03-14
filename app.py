import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from PIL import ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD

from converter_core import (
    EXPORT_HANDLERS,
    default_extension_for,
    render_pdf_full,
    render_pdf_preview,
    suggested_filetypes,
)
import updater


ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class PDFHandlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDFHandler")
        self.root.geometry("1440x860")
        self.root.minsize(1150, 720)

        self.pdf_path = None
        self.preview_pages = []
        self.converted_pages = []

        self.left_image_refs = []
        self.right_image_refs = []

        self.queue = queue.Queue()
        self.conversion_running = False

        self.dpi_var = tk.StringVar(value="450")
        self.status_var = tk.StringVar(value="Drop a PDF here or click Import PDF")
        self.progress_var = tk.DoubleVar(value=0)

        self.build_ui()
        self.root.after(100, self.poll_queue)

    def build_ui(self):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        # Top bar
        topbar = ctk.CTkFrame(self.root, corner_radius=0, height=68)
        topbar.grid(row=0, column=0, sticky="ew")
        topbar.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            topbar,
            text="PDFHandler",
            font=ctk.CTkFont(size=26, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=16, sticky="w")

        ctk.CTkLabel(
            topbar,
            text="PDF to Slides — clean, sharp, painless",
            text_color="#9aa0aa",
            font=ctk.CTkFont(size=13),
        ).grid(row=0, column=0, padx=170, pady=20, sticky="w")

        ctk.CTkButton(
            topbar,
            text="Check for Updates",
            width=160,
            command=self.check_updates,
        ).grid(row=0, column=1, padx=16, pady=14)

        # Main
        main = ctk.CTkFrame(self.root, corner_radius=0)
        main.grid(row=1, column=0, sticky="nsew", padx=14, pady=14)
        main.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=1)
        main.grid_rowconfigure(2, weight=1)

        # Left header
        left_header = ctk.CTkFrame(main)
        left_header.grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=(0, 10))
        left_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            left_header,
            text="Source Preview",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, padx=14, pady=12, sticky="w")

        ctk.CTkButton(
            left_header,
            text="Import PDF",
            width=120,
            command=self.pick_pdf,
        ).grid(row=0, column=1, padx=12, pady=12)

        # Right header
        right_header = ctk.CTkFrame(main)
        right_header.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=(0, 10))
        right_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            right_header,
            text="Converted Preview",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, padx=14, pady=12, sticky="w")

        self.export_btn = ctk.CTkButton(
            right_header,
            text="Convert / Export",
            width=150,
            state="disabled",
            command=self.open_export_menu,
        )
        self.export_btn.grid(row=0, column=1, padx=12, pady=12)

        # Drop zone
        self.dropzone = ctk.CTkFrame(main, height=140)
        self.dropzone.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(0, 10))
        self.dropzone.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.dropzone,
            text="Drop PDF Here",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).grid(row=0, column=0, pady=(22, 4), padx=10)

        ctk.CTkLabel(
            self.dropzone,
            text="or click Import PDF",
            font=ctk.CTkFont(size=14),
            text_color="#aab2c0",
        ).grid(row=1, column=0, pady=(0, 20), padx=10)

        self.dropzone.drop_target_register(DND_FILES)
        self.dropzone.dnd_bind("<<Drop>>", self.handle_drop)

        # Panels
        self.left_scroll = ctk.CTkScrollableFrame(main, label_text="Slides")
        self.left_scroll.grid(row=2, column=0, sticky="nsew", padx=(0, 8), pady=(0, 0))

        self.right_scroll = ctk.CTkScrollableFrame(main, label_text="Rendered Output")
        self.right_scroll.grid(row=1, column=1, rowspan=2, sticky="nsew", padx=(8, 0), pady=(0, 0))

        # Bottom bar
        bottom = ctk.CTkFrame(self.root)
        bottom.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 14))
        bottom.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(
            bottom,
            text="Output DPI",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, padx=(16, 8), pady=12)

        ctk.CTkSegmentedButton(
            bottom,
            values=["300", "450", "600"],
            variable=self.dpi_var,
        ).grid(row=0, column=1, padx=8, pady=12, sticky="w")

        self.progress = ctk.CTkProgressBar(bottom, variable=self.progress_var)
        self.progress.grid(row=0, column=2, padx=16, pady=12, sticky="ew")
        self.progress.set(0)

        ctk.CTkLabel(
            bottom,
            textvariable=self.status_var,
            text_color="#c2c9d2",
            anchor="e",
        ).grid(row=0, column=3, padx=16, pady=12, sticky="e")

    def pick_pdf(self):
        path = filedialog.askopenfilename(
            title="Select a PDF",
            filetypes=[("PDF Files", "*.pdf")],
        )
        if path:
            self.load_pdf(path)

    def handle_drop(self, event):
        try:
            files = self.root.tk.splitlist(event.data)
            if not files:
                return
            path = files[0].strip()
            if not path.lower().endswith(".pdf"):
                messagebox.showerror("Not a PDF", "Please drop a PDF file.")
                return
            self.load_pdf(path)
        except Exception as exc:
            messagebox.showerror("Drop Error", str(exc))

    def load_pdf(self, path):
        self.pdf_path = path
        self.preview_pages = []
        self.converted_pages = []
        self.clear_scroll_frame(self.left_scroll)
        self.clear_scroll_frame(self.right_scroll)
        self.left_image_refs.clear()
        self.right_image_refs.clear()
        self.progress.set(0)
        self.status_var.set("Rendering source preview...")

        try:
            self.preview_pages = render_pdf_preview(path, preview_dpi=110)
            self.populate_preview_panel(self.left_scroll, self.preview_pages, self.left_image_refs)
            self.status_var.set(f"Loaded: {os.path.basename(path)}")
            self.export_btn.configure(state="normal")
        except Exception as exc:
            self.status_var.set("Failed to load PDF")
            messagebox.showerror("Preview Error", str(exc))

    def populate_preview_panel(self, parent, pages, image_ref_list):
        self.clear_scroll_frame(parent)
        image_ref_list.clear()

        for page in pages:
            card = ctk.CTkFrame(parent)
            card.pack(fill="x", padx=8, pady=8)

            thumb = page.image.copy()
            thumb.thumbnail((390, 520))
            photo = ImageTk.PhotoImage(thumb)
            image_ref_list.append(photo)

            img_label = tk.Label(card, image=photo, bd=0)
            img_label.pack(padx=10, pady=(10, 6))

            ctk.CTkLabel(
                card,
                text=f"Page {page.page_number}",
                font=ctk.CTkFont(size=14, weight="bold"),
            ).pack(pady=(0, 8))

    def add_converted_preview_card(self, page):
        card = ctk.CTkFrame(self.right_scroll)
        card.pack(fill="x", padx=8, pady=8)

        thumb = page.image.copy()
        thumb.thumbnail((390, 520))
        photo = ImageTk.PhotoImage(thumb)
        self.right_image_refs.append(photo)

        img_label = tk.Label(card, image=photo, bd=0)
        img_label.pack(padx=10, pady=(10, 6))

        ctk.CTkLabel(
            card,
            text=f"Rendered Page {page.page_number}",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(pady=(0, 8))

    def clear_scroll_frame(self, frame):
        for child in frame.winfo_children():
            child.destroy()

    def open_export_menu(self):
        if not self.pdf_path:
            messagebox.showwarning("No PDF", "Please import a PDF first.")
            return
        if self.conversion_running:
            return

        top = ctk.CTkToplevel(self.root)
        top.title("Choose Export Format")
        top.geometry("700x430")
        top.transient(self.root)
        top.grab_set()

        ctk.CTkLabel(
            top,
            text="Export As",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).pack(pady=(18, 4))

        ctk.CTkLabel(
            top,
            text="Pick a slideshow-friendly format.",
            text_color="#aab2c0",
        ).pack(pady=(0, 14))

        grid = ctk.CTkFrame(top)
        grid.pack(fill="both", expand=True, padx=18, pady=12)

        formats = [
            ("pptx", "PowerPoint (.pptx)", "Best for PowerPoint and general slide use."),
            ("gslides_pptx", "Google Slides Compatible (.pptx)", "Exports PPTX. Upload it to Google Drive and open in Google Slides."),
            ("png_zip", "PNG Slide Set (.zip)", "One PNG image per slide."),
            ("jpg_zip", "JPG Slide Set (.zip)", "One JPG image per slide, smaller files."),
        ]

        for idx, (fmt, title, desc) in enumerate(formats):
            card = ctk.CTkFrame(grid)
            r, c = divmod(idx, 2)
            card.grid(row=r, column=c, padx=10, pady=10, sticky="nsew")
            grid.grid_columnconfigure(c, weight=1)
            grid.grid_rowconfigure(r, weight=1)

            ctk.CTkLabel(
                card,
                text=title,
                font=ctk.CTkFont(size=16, weight="bold"),
                wraplength=250,
                justify="left",
            ).pack(anchor="w", padx=12, pady=(12, 4))

            ctk.CTkLabel(
                card,
                text=desc,
                text_color="#aab2c0",
                wraplength=250,
                justify="left",
            ).pack(anchor="w", padx=12, pady=(0, 12))

            ctk.CTkButton(
                card,
                text="Choose",
                command=lambda f=fmt, win=top: self.start_export(f, win),
            ).pack(anchor="e", padx=12, pady=(0, 12))

    def start_export(self, fmt, window):
        ext = default_extension_for(fmt)
        types = suggested_filetypes(fmt)

        base = "output"
        if self.pdf_path:
            base = os.path.splitext(os.path.basename(self.pdf_path))[0]

        out_path = filedialog.asksaveasfilename(
            title="Save Export",
            defaultextension=ext,
            filetypes=types,
            initialfile=base + ext,
        )
        if not out_path:
            return

        window.destroy()

        self.converted_pages = []
        self.right_image_refs.clear()
        self.clear_scroll_frame(self.right_scroll)
        self.progress.set(0)
        self.conversion_running = True
        self.export_btn.configure(state="disabled")
        self.status_var.set(f"Rendering at {self.dpi_var.get()} DPI...")

        worker = threading.Thread(
            target=self._convert_worker,
            args=(fmt, out_path, int(self.dpi_var.get())),
            daemon=True,
        )
        worker.start()

    def _convert_worker(self, fmt, out_path, dpi):
        try:
            rendered_pages = []

            def on_progress(current, total, page):
                rendered_pages.append(page)
                self.queue.put(("page", page, current, total))

            render_pdf_full(self.pdf_path, dpi=dpi, progress_callback=on_progress)

            exporter = EXPORT_HANDLERS[fmt]
            exporter(rendered_pages, out_path)

            self.queue.put(("done", out_path))
        except Exception as exc:
            self.queue.put(("error", str(exc)))

    def poll_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]

                if kind == "page":
                    _, page, current, total = item
                    self.converted_pages.append(page)
                    self.add_converted_preview_card(page)
                    self.progress.set(current / total)
                    self.status_var.set(f"Rendering page {current} of {total}...")

                elif kind == "done":
                    out_path = item[1]
                    self.conversion_running = False
                    self.export_btn.configure(state="normal")
                    self.progress.set(1)
                    self.status_var.set(f"Finished: {os.path.basename(out_path)}")
                    self.show_done_popup(out_path)

                elif kind == "error":
                    self.conversion_running = False
                    self.export_btn.configure(state="normal")
                    self.status_var.set("Conversion failed")
                    messagebox.showerror("Export Error", item[1])

        except queue.Empty:
            pass

        self.root.after(100, self.poll_queue)

    def show_done_popup(self, out_path):
        top = ctk.CTkToplevel(self.root)
        top.title("Export Complete")
        top.geometry("430x210")
        top.transient(self.root)
        top.grab_set()

        ctk.CTkLabel(
            top,
            text="Export Complete",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(pady=(18, 8))

        ctk.CTkLabel(
            top,
            text=os.path.basename(out_path),
            wraplength=360,
            text_color="#aab2c0",
        ).pack(pady=(0, 14))

        btns = ctk.CTkFrame(top)
        btns.pack(pady=10)

        ctk.CTkButton(
            btns,
            text="Open File",
            command=lambda: os.startfile(out_path),
        ).grid(row=0, column=0, padx=8, pady=8)

        ctk.CTkButton(
            btns,
            text="Open Folder",
            command=lambda: os.startfile(os.path.dirname(out_path)),
        ).grid(row=0, column=1, padx=8, pady=8)

        ctk.CTkButton(
            btns,
            text="Close",
            fg_color="gray25",
            command=top.destroy,
        ).grid(row=0, column=2, padx=8, pady=8)

    def check_updates(self):
        result = updater.check_for_updates()

        if result is None:
            messagebox.showinfo(
                "Updates Not Configured",
                "Update checking is not configured yet.\n\nAdd VERSION_URL and DOWNLOAD_URL in updater.py later."
            )
            return

        if not result.get("ok"):
            messagebox.showerror("Update Check Failed", result.get("message", "Unknown error"))
            return

        if result.get("has_update"):
            if messagebox.askyesno(
                "Update Available",
                f"Version {result['latest']} is available.\n\nOpen download page?"
            ):
                updater.open_download_page(result.get("download_url", ""))
        else:
            messagebox.showinfo(
                "Up to Date",
                f"You already have the latest version ({updater.CURRENT_VERSION})."
            )


def main():
    root = TkinterDnD.Tk()
    app = PDFHandlerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()