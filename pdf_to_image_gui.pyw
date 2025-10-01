# pdf_to_image_gui.py
# ------------------------------------------------------------
# PDFを「ページ単位」で画像に一括変換するGUIツール
# - 依存: PyMuPDF (fitz)
# - 使い方:
#   1) 「PDFを選ぶ」から複数PDFを選択
#   2) 「書き出し先フォルダ」を選択
#   3) DPIを指定（既定: 300）または「PDFネイティブ解像度(72dpi)で描画」にチェック
#   4) 「変換スタート」で各ページをPNG/JPGに保存
# - 仕様:
#   * 各PDFごとに出力先/ファイル名を自動生成
#   * 拡張子はPNG/JPEGから選択（既定: PNG）
#   * 進捗ログ表示
#   * 単一ファイルのみ出力するオプションあり（サブフォルダ作成なし）
# ------------------------------------------------------------

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import fitz  # PyMuPDF
except Exception as e:
    raise SystemExit("PyMuPDF(fitz) が見つかりません。先に 'pip install PyMuPDF' を実行してください")

APP_TITLE = "PDF → 画像 変換ツール"
APP_VERSION = "v1.1"

class PdfToImageGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("720x520")
        self.minsize(640, 480)

        self.pdf_paths = []
        self.output_dir = ""

        self.var_dpi = tk.StringVar(value="300")
        self.var_native = tk.BooleanVar(value=False)
        self.var_format = tk.StringVar(value="PNG")  # PNG / JPEG
        self.var_subfolder = tk.BooleanVar(value=False)  # ←デフォルト: サブフォルダなし（※複数PDF選択時は自動でON）
        self.var_quality = tk.IntVar(value=95)  # JPEG品質

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # ファイル選択
        frame_files = ttk.LabelFrame(self, text="入力PDF")
        frame_files.pack(fill="x", **pad)

        btn_pick = ttk.Button(frame_files, text="PDFを選ぶ (複数OK)", command=self.pick_pdfs)
        btn_pick.pack(side="left", padx=10, pady=8)

        self.lbl_count = ttk.Label(frame_files, text="未選択")
        self.lbl_count.pack(side="left")

        # 出力先
        frame_out = ttk.LabelFrame(self, text="書き出し先")
        frame_out.pack(fill="x", **pad)

        btn_out = ttk.Button(frame_out, text="書き出し先フォルダ", command=self.pick_output)
        btn_out.pack(side="left", padx=10, pady=8)

        self.lbl_out = ttk.Label(frame_out, text="未選択")
        self.lbl_out.pack(side="left")

        # オプション
        frame_opts = ttk.LabelFrame(self, text="オプション")
        frame_opts.pack(fill="x", **pad)

        row1 = ttk.Frame(frame_opts)
        row1.pack(fill="x", padx=8, pady=4)
        ttk.Label(row1, text="DPI:").pack(side="left")
        self.entry_dpi = ttk.Entry(row1, textvariable=self.var_dpi, width=8)
        self.entry_dpi.pack(side="left")
        ttk.Label(row1, text="(数値が大きいほど高解像度/ファイル大)").pack(side="left", padx=6)

        chk_native = ttk.Checkbutton(
            frame_opts,
            text="PDFネイティブ解像度(72dpi)で描画 (拡大/縮小しない)",
            variable=self.var_native,
            command=self._toggle_dpi_entry,
        )
        chk_native.pack(anchor="w", padx=8)

        row2 = ttk.Frame(frame_opts)
        row2.pack(fill="x", padx=8, pady=4)
        ttk.Label(row2, text="画像形式:").pack(side="left")
        fmt = ttk.Combobox(row2, textvariable=self.var_format, values=["PNG", "JPEG"], width=7, state="readonly")
        fmt.pack(side="left")

        ttk.Label(row2, text="  JPEG品質(1-100):").pack(side="left")
        quality = ttk.Spinbox(row2, from_=1, to=100, textvariable=self.var_quality, width=5)
        quality.pack(side="left")

        chk_sub = ttk.Checkbutton(frame_opts, text="PDFごとにサブフォルダを作る", variable=self.var_subfolder)
        chk_sub.pack(anchor="w", padx=8)

        # 実行
        frame_run = ttk.Frame(self)
        frame_run.pack(fill="x", **pad)
        self.btn_run = ttk.Button(frame_run, text="変換スタート", command=self.run_convert)
        self.btn_run.pack(side="left", padx=10)

        self.progress = ttk.Progressbar(frame_run, orient="horizontal", mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)

        # ログ
        frame_log = ttk.LabelFrame(self, text="ログ")
        frame_log.pack(fill="both", expand=True, **pad)
        self.txt_log = tk.Text(frame_log, height=12, wrap="none")
        self.txt_log.pack(fill="both", expand=True, padx=8, pady=8)

        self._toggle_dpi_entry()

    def _toggle_dpi_entry(self):
        state = "disabled" if self.var_native.get() else "normal"
        self.entry_dpi.configure(state=state)

    def pick_pdfs(self):
        paths = filedialog.askopenfilenames(
            title="PDFファイルを選択",
            filetypes=[("PDF", "*.pdf")],
        )
        if not paths:
            return
        self.pdf_paths = list(paths)
        self.lbl_count.configure(text=f"{len(self.pdf_paths)} 件選択")
        # 複数PDFならサブフォルダを自動ON（名前衝突回避 & 整理のため）
        if len(self.pdf_paths) > 1:
            self.var_subfolder.set(True)
        self.log(f"PDF選択: {len(self.pdf_paths)}件")

    def pick_output(self):
        d = filedialog.askdirectory(title="書き出し先フォルダ")
        if not d:
            return
        self.output_dir = d
        self.lbl_out.configure(text=self.output_dir)
        self.log(f"出力先: {self.output_dir}")

    def run_convert(self):
        if not self.pdf_paths:
            messagebox.showwarning("入力エラー", "PDFが選ばれていません")
            return
        if not self.output_dir:
            messagebox.showwarning("出力エラー", "書き出し先フォルダを選んでください")
            return

        # DPI設定
        if self.var_native.get():
            dpi = 72
        else:
            try:
                dpi = int(self.var_dpi.get())
                if dpi <= 0:
                    raise ValueError
            except Exception:
                messagebox.showerror("DPIエラー", "DPIは正の整数で入力してください")
                return

        fmt = self.var_format.get().upper()
        make_sub = self.var_subfolder.get()
        quality = int(self.var_quality.get()) if fmt == "JPEG" else None

        # 進捗最大値: 総ページ数
        total_pages = 0
        for p in self.pdf_paths:
            try:
                with fitz.open(p) as doc:
                    total_pages += doc.page_count
            except Exception as e:
                self.log(f"[SKIP] 開けませんでした: {p}  ({e})")
        if total_pages == 0:
            messagebox.showerror("エラー", "有効なPDFがありません")
            return

        self.progress.configure(maximum=total_pages, value=0)
        self.btn_run.configure(state="disabled")
        self.update_idletasks()

        pages_done = 0
        for pdf_path in self.pdf_paths:
            try:
                base = os.path.splitext(os.path.basename(pdf_path))[0]
                out_dir = self.output_dir if not make_sub else os.path.join(self.output_dir, base)
                if make_sub:
                    os.makedirs(out_dir, exist_ok=True)

                with fitz.open(pdf_path) as doc:
                    zoom = dpi / 72.0
                    mat = fitz.Matrix(zoom, zoom)

                    self.log(f"— 変換開始: {os.path.basename(pdf_path)} (ページ数: {doc.page_count}, dpi={dpi}, fmt={fmt})")

                    for i, page in enumerate(doc, start=1):
                        pix = page.get_pixmap(matrix=mat, alpha=False)
                        out_name = f"{base}_p{i:03d}.{ 'png' if fmt=='PNG' else 'jpg' }"
                        out_path = os.path.join(out_dir, out_name)

                        if fmt == "PNG":
                            pix.save(out_path)
                        else:
                            pix.save(out_path, jpg_quality=quality)

                        pages_done += 1
                        self.progress.configure(value=pages_done)
                        if pages_done % 5 == 0:
                            self.update_idletasks()

                    self.log(f"✓ 完了: {os.path.basename(pdf_path)} → {out_dir}")

            except Exception as e:
                self.log(f"[ERROR] {pdf_path}: {e}")

        self.btn_run.configure(state="normal")
        messagebox.showinfo("完了", "すべての変換が完了しました！")

    def log(self, msg: str):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")


def main():
    app = PdfToImageGUI()
    app.mainloop()


if __name__ == "__main__":
    main()