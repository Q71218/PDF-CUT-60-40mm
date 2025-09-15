import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import io, os, sys

# ========== 常數 ==========
def mm_to_pt(mm: float) -> float:
    """毫米轉換為 PDF points (1pt = 1/72 inch)."""
    return mm * 72 / 25.4

PAGE_W = mm_to_pt(60)   # 60 mm
PAGE_H = mm_to_pt(40)   # 40 mm

class PDFCropper:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 裁切工具    YGA設計")

        self.pdf_doc = None
        self.pdf_path = None
        self.current_page = 0
        self.zoom = 1.0
        self.image = None
        self.tk_img = None
        self.selections = []  # (page_index, fitz.Rect, canvas_rect_id)

        # 說明文字
        self.info_label = tk.Label(
            root,
            text="簡易說明：👉 使用滑鼠自由框選，右鍵可取消當下框選",
            bg="lightyellow",
            fg="black",
            anchor="w"
        )
        self.info_label.pack(side="top", fill="x")

        # 工具列
        self._build_toolbar()

        # 畫布
        self.canvas = tk.Canvas(root, bg="gray")
        self.canvas.pack(fill="both", expand=True)

        # 狀態列
        self.status_label = tk.Label(
            root, text="尚未載入 PDF", anchor="w", relief="sunken"
        )
        self.status_label.pack(side="bottom", fill="x")

        # 滑鼠事件
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.canvas.bind("<Control-MouseWheel>", self.on_ctrl_scroll)

        self.start_x = self.start_y = None
        self.rect_id = None

    def _build_toolbar(self):
        toolbar = ttk.Frame(self.root)
        toolbar.pack(side="top", fill="x")
        ttk.Button(toolbar, text="開啟 PDF", command=self.open_pdf).pack(side="left", padx=2, pady=2)
        ttk.Button(toolbar, text="上一頁", command=self.prev_page).pack(side="left", padx=2, pady=2)
        ttk.Button(toolbar, text="下一頁", command=self.next_page).pack(side="left", padx=2, pady=2)
        ttk.Button(toolbar, text="縮小", command=lambda: self.change_zoom(0.8)).pack(side="left", padx=2, pady=2)
        ttk.Button(toolbar, text="放大", command=lambda: self.change_zoom(1.25)).pack(side="left", padx=2, pady=2)
        ttk.Button(toolbar, text="匯出 PDF", command=self.export_pdf).pack(side="left", padx=2, pady=2)

    def update_status(self):
        """更新狀態列顯示"""
        if self.pdf_doc:
            total = len(self.pdf_doc)
            zoom_percent = int(self.zoom * 100)
            selection_count = len(self.selections)
            self.status_label.config(
                text=f"頁面: {self.current_page+1}/{total} | 縮放: {zoom_percent}% | 框選數量: {selection_count}"
            )
        else:
            self.status_label.config(text="尚未載入 PDF")

    def open_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not path:
            return
        self.pdf_path = path
        self.pdf_doc = fitz.open(path)
        self.current_page = 0
        self.selections.clear()
        self.zoom = 1.0
        self.show_page()

    def show_page(self):
        if not self.pdf_doc:
            return
        page = self.pdf_doc[self.current_page]
        pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom))  # 預覽用點陣圖
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        self.image = img
        self.tk_img = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)

        for pgno, rect, rid in self.selections:
            if pgno == self.current_page:
                x0, y0 = rect.x0 * self.zoom, rect.y0 * self.zoom
                x1, y1 = rect.x1 * self.zoom, rect.y1 * self.zoom
                new_id = self.canvas.create_rectangle(x0, y0, x1, y1, outline="red", width=2)
                self._update_rect_id(pgno, rect, new_id)

        self.update_status()

    def prev_page(self):
        if self.pdf_doc and self.current_page > 0:
            self.current_page -= 1
            self.show_page()

    def next_page(self):
        if self.pdf_doc and self.current_page < len(self.pdf_doc) - 1:
            self.current_page += 1
            self.show_page()

    def change_zoom(self, factor: float):
        if self.pdf_doc:
            self.zoom *= factor
            self.show_page()

    def on_press(self, event):
        self.start_x, self.start_y = event.x, event.y
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(
            event.x, event.y, event.x, event.y, outline="red", width=2
        )

    def on_drag(self, event):
        if self.rect_id:
            self.canvas.coords(self.rect_id, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        if not self.image or not self.pdf_doc:
            return
        x0, y0, x1, y1 = self.canvas.coords(self.rect_id)
        scale = 1 / self.zoom
        rect = fitz.Rect(x0 * scale, y0 * scale, x1 * scale, y1 * scale)
        self.selections.append((self.current_page, rect, self.rect_id))
        self.rect_id = None
        self.update_status()

    def on_right_click(self, event):
        to_delete = None
        for pgno, rect, rid in self.selections:
            if pgno == self.current_page:
                x0, y0, x1, y1 = self.canvas.coords(rid)
                if x0 <= event.x <= x1 and y0 <= event.y <= y1:
                    to_delete = (pgno, rect, rid)
                    break
        if to_delete:
            self.canvas.delete(to_delete[2])
            self.selections.remove(to_delete)
            self.update_status()

    def on_ctrl_scroll(self, event):
        if event.delta > 0:
            self.change_zoom(1.1)
        else:
            self.change_zoom(0.9)

    def _update_rect_id(self, pgno, rect, new_id):
        for i, (p, r, rid) in enumerate(self.selections):
            if p == pgno and r == rect:
                self.selections[i] = (p, r, new_id)
                break

    def export_pdf(self):
        if not self.selections:
            messagebox.showwarning("警告", "沒有框選區域！")
            return

        out_doc = fitz.open()
        for pgno, rect, _ in self.selections:
            new_page = out_doc.new_page(width=PAGE_W, height=PAGE_H)
            new_page.show_pdf_page(
                fitz.Rect(0, 0, PAGE_W, PAGE_H),
                self.pdf_doc,
                pgno,
                clip=rect
            )

        base, _ = os.path.splitext(self.pdf_path)
        outname = f"{base}_{len(self.selections)}.pdf"
        out_doc.save(outname)
        out_doc.close()
        messagebox.showinfo("完成", f"已輸出：{outname}")
        self.update_status()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCropper(root)
    root.mainloop()
