import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import json
from pathlib import Path
import backend

SETTINGS_FILE = "settings.json"

def main():
    root = tk.Tk()
    app = ImageDocxGUI(root)
    root.mainloop()

class ImageDocxGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Chèn ảnh vào file Word tự động")

        self.settings = self.load_settings()

        self.input_docs = tk.StringVar(value=self.settings.get("input_docs", ""))
        self.input_imgs = tk.StringVar(value=self.settings.get("input_imgs", ""))
        self.output_docs = tk.StringVar(value=self.settings.get("output_docs", ""))

        self.make_row("Thư mục file docx đầu vào", self.input_docs)
        self.make_row("Thư mục file ảnh đầu vào", self.input_imgs)
        self.make_row("Thư mục file docx đầu ra", self.output_docs)

        run_btn = tk.Button(root, text="Chạy", command=self.run_process, bg="#4CAF50", fg="white")
        run_btn.pack(pady=10)

        self.log = scrolledtext.ScrolledText(root, height=15, width=80, state="disabled")
        self.log.pack(padx=10, pady=10)

    def make_row(self, label, var):
        frame = tk.Frame(self.root)
        frame.pack(fill="x", padx=10, pady=5)
        tk.Label(frame, text=label, width=20, anchor="w").pack(side="left")
        tk.Entry(frame, textvariable=var, width=50).pack(side="left", padx=5)
        tk.Button(frame, text="Tìm kiếm", command=lambda: self.browse_folder(var)).pack(side="left")

    def browse_folder(self, var):
        folder = filedialog.askdirectory()
        if folder:
            var.set(folder)

    def log_write(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")
        self.root.update_idletasks()

    def run_process(self):
        docs = Path(self.input_docs.get())
        imgs = Path(self.input_imgs.get())
        out = Path(self.output_docs.get())

        if not docs.exists() or not imgs.exists() or not out.exists():
            messagebox.showerror("Lỗi", "Hãy chọn thư mục hợp lệ")
            return

        self.settings["input_docs"] = str(docs)
        self.settings["input_imgs"] = str(imgs)
        self.settings["output_docs"] = str(out)
        self.save_settings()

        self.log_write("=== Bắt đầu chạy ===")

        skipped, processed = backend.run_batch(docs, imgs, out, log_func=self.log_write)

        self.log_write(f"\n=== TỔNG HỢP ===")
        self.log_write(f"Số file đã tổng hợp: {processed}")
        if skipped:
            self.log_write(f"Bỏ qua (những) file: {len(skipped)}")
            for fname, reason in skipped:
                self.log_write(f"  - {fname}: {reason}")
        else:
            self.log_write("Đã tổng hợp tất cả các file.")

    def load_settings(self):
        if Path(SETTINGS_FILE).exists():
            return json.load(open(SETTINGS_FILE, "r", encoding="utf-8"))
        return {}

    def save_settings(self):
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(self.settings, f, indent=2, ensure_ascii=False)

if __name__ == "__main__":
    main()