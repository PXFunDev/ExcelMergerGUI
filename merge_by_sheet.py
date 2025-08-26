"""
============================================
Excelファイル結合ツール
============================================

プログラム名: merge_by_sheet.py

作成者  : PSC 平野 翔士
作成日  : 2025年8月25日
説明    : このプログラムは、Excelファイルを同名シートごとに縦結合するためのものです。
使用方法: プログラムを実行

最終更新日: 2025年8月25日
最終更新者: PSC 平野 翔士

変更履歴:
日付        バージョン    変更内容
----------  ----------  -----------------------------------
2025/08/25  1.0         初版作成

--------------------------------------------

## 要件定義
- GUIにドラッグアンドドロップしたファイルを結合
- GUIに取り込んだファイルを並べ替えることができる
- GUIに取り込んだファイルを削除することができる
    - 選択したファイルを削除
    - すべてのファイルを削除
- 設定項目
    - 結合シート名（ホワイトリスト）
    - ヘッダーの行数設定
    - 結合の出力先を指定できる
- 実行したら同一シート名ごとに結合したExcelファイルを出力する
"""

"""
ExcelファイルをGUIで選択/ドラッグ&ドロップし、
同名シートごとに縦結合して出力するツール。

主な機能:
- ファイルのドラッグ&ドロップ（tkinterdnd2）
- ファイルリストの順序変更（↑/↓）
- 選択削除/全削除
- 結合対象シートのホワイトリスト（カンマ区切り）
- ヘッダー行番号の指定（0=1行目がヘッダー）
- 保存先の指定
- 同名シートごとに列ユニオンで縦結合（欠損は空欄）
- 追跡列 _source_file, _source_sheet を付与
"""

import re
from typing import Any
import pandas as pd
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD


# --- tkinterdnd2 が無くても起動できるようにする（DnDなしで動作） ---
DND_AVAILABLE = True
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_AVAILABLE = False


# -------------------------------
# ユーティリティ関数
# -------------------------------
def is_excel_file(path: str) -> bool:
    """対応拡張子チェック（基本は xlsx/xlsm/xls だけに絞るのが安定）"""
    return str(path).lower().endswith((".xlsx", ".xlsm", ".xls"))

def parse_dnd_paths(dnd_payload: str) -> list[str]:
    """
    DnDで受け取る文字列をファイルパス配列へ。
    スペース・日本語対応のため {} やクォートを外しつつ分割。
    """
    # {C:/aa bb/ccc.xlsx} {D:/テスト.xlsx}
    parts = re.findall(r"\{.*?\}|[^\s]+", dnd_payload)
    clean = [p.strip("{}") for p in parts]
    return clean


# -------------------------------
# 結合ロジック（テストしやすいようGUIから分離）
# -------------------------------
def merge_by_sheet(files: list[Path], sheet_whitelist: list[str], header_row: int, out_path: Path) -> None:
    """
    指定ファイル群を、同名シートごとに縦結合して out_path に保存する。
    - files: 結合対象Excel（順序はこのまま維持）
    - sheet_whitelist: 指定があれば、そのシート名のみ結合（空なら全シート）
    - header_row: pandas.read_excel の header に渡す値（0=1行目をヘッダー）
    - out_path: 出力先 .xlsx
    """
    buckets: dict[str, list[Any]] = {}

    for f in files:
        try:
            with pd.ExcelFile(f) as xf:
                target_sheets = xf.sheet_names
                if sheet_whitelist:
                    # ホワイトリスト（完全一致）。trimして比較
                    wl = [s.strip() for s in sheet_whitelist if s.strip()]
                    target_sheets = [s for s in xf.sheet_names if s in wl]

                for sheet in target_sheets:
                    try:
                        # 文字列で読み込み → 型ブレ最小化（必要ならあとで変換）
                        df = pd.read_excel(xf, sheet_name=sheet, header=header_row, dtype=str)
                        if df.empty:
                            continue
                        # 余計な Unnamed: 列は削除（よくある空白列）
                        df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
                        # 追跡列を付与
                        df["_source_file"] = f.name
                        df["_source_sheet"] = sheet
                        buckets.setdefault(str(sheet), []).append(df)
                    except Exception as e:
                        print(f"[WARN] 読み込み失敗: {f.name} / {sheet}: {e}")
                        continue
        except Exception as e:
            print(f"[WARN] 開けませんでした: {f.name}: {e}")
            continue

    if not buckets:
        raise RuntimeError("どのシートも読み込めませんでした。ホワイトリスト・ヘッダー設定を確認してください。")

    # 書き出し
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name, df_list in buckets.items():
            try:
                merged = pd.concat(df_list, ignore_index=True, sort=False)  # 列ユニオン
                # Excelのシート名31文字制限
                safe_name = sheet_name[:31] if len(sheet_name) > 31 else sheet_name
                merged.to_excel(writer, sheet_name=safe_name, index=False)
            except Exception as e:
                print(f"[WARN] 書き出し失敗: {sheet_name}: {e}")
                continue


# -------------------------------
# GUI 本体
# -------------------------------
class ExcelMergerGUI:
    """画面・イベント・結合起動を司るクラス"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel Merger - 同名シート縦結合")
        self.root.geometry("820x520")

        # ファイル格納（順序を保持）
        self.files: list[Path] = []

        self._build_widgets()

    # ---- 画面構築 ----
    def _build_widgets(self):
        pad: dict[str, Any] = {"padx": 8, "pady": 6}

        # 上段：ファイル操作
        frm_files = ttk.LabelFrame(self.root, text="ファイル")
        frm_files.pack(fill="both", expand=False, **pad) 

        btn_add = ttk.Button(frm_files, text="ファイルを追加", command=self.on_add_files)
        btn_add.grid(row=0, column=0, **pad) 

        btn_up = ttk.Button(frm_files, text="↑ 上へ", command=self.on_move_up)
        btn_up.grid(row=0, column=1, **pad) 

        btn_down = ttk.Button(frm_files, text="↓ 下へ", command=self.on_move_down)
        btn_down.grid(row=0, column=2, **pad) 

        btn_del = ttk.Button(frm_files, text="選択削除", command=self.on_delete_selected)
        btn_del.grid(row=0, column=3, **pad) 

        btn_clear = ttk.Button(frm_files, text="全削除", command=self.on_clear)
        btn_clear.grid(row=0, column=4, **pad) 

        # DnDヒント
        lbl_hint = ttk.Label(frm_files, text="ここにドラッグ＆ドロップでも追加できます")
        lbl_hint.grid(row=0, column=5, sticky="w", **pad) 

        # ファイルリスト
        self.lst = tk.Listbox(frm_files, selectmode=tk.EXTENDED, height=8)
        self.lst.grid(row=1, column=0, columnspan=6, sticky="nsew", **pad) 
        frm_files.grid_columnconfigure(5, weight=1)
        frm_files.grid_rowconfigure(1, weight=1)

        # DnD 対応（tkinterdnd2 があれば）
        if DND_AVAILABLE and isinstance(self.root, TkinterDnD.Tk):
            self.lst.drop_target_register(DND_FILES) # type: ignore
            self.lst.dnd_bind("<<Drop>>", self.on_drop) # type: ignore
        else:
            lbl_hint.configure(text="（DnD無効: tkinterdnd2 が未導入 or 非対応環境）")

        # 中段：設定
        frm_opts = ttk.LabelFrame(self.root, text="設定")
        frm_opts.pack(fill="x", expand=False, **pad) 

        ttk.Label(frm_opts, text="結合シート名（カンマ区切り。空なら全シート）").grid(row=0, column=0, sticky="w", **pad)
        self.ent_whitelist = ttk.Entry(frm_opts, width=60)
        self.ent_whitelist.grid(row=0, column=1, columnspan=3, sticky="we", **pad)

        ttk.Label(frm_opts, text="ヘッダー行番号（0=1行目）").grid(row=1, column=0, sticky="w", **pad)
        self.ent_header = ttk.Entry(frm_opts, width=10)
        self.ent_header.insert(0, "0")
        self.ent_header.grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(frm_opts, text="保存先").grid(row=2, column=0, sticky="w", **pad)
        self.out_var = tk.StringVar()
        self.ent_out = ttk.Entry(frm_opts, textvariable=self.out_var, width=60)
        self.ent_out.grid(row=2, column=1, sticky="we", **pad)

        btn_browse = ttk.Button(frm_opts, text="参照…", command=self.on_browse_out)
        btn_browse.grid(row=2, column=2, **pad)

        frm_opts.grid_columnconfigure(1, weight=1)

        # 下段：実行
        frm_run = ttk.Frame(self.root)
        frm_run.pack(fill="x", expand=False, **pad)

        self.btn_run = ttk.Button(frm_run, text="結合を実行", command=self.on_run)
        self.btn_run.pack(side="right")

        # 既定の保存名をセット
        self._set_default_outname()

    # ---- イベント群 ----
    def on_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Excelファイルを選択",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]
        )
        self._add_paths(paths)

    def on_drop(self, event):
        paths = parse_dnd_paths(event.data)
        self._add_paths(paths)

    def _add_paths(self, paths):
        added = 0
        for p in paths:
            if is_excel_file(p):
                path = Path(p)
                if path.exists() and path not in self.files:
                    self.files.append(path)
                    self.lst.insert(tk.END, str(path))
                    added += 1
        if added == 0 and paths:
            messagebox.showwarning("警告", "Excelファイル（.xlsx/.xlsm/.xls）のみ追加できます。")

    def on_move_up(self):
        # 選択行を一つ上へ
        sel = list(self.lst.curselection())
        if not sel:
            return
        for i in sel:
            if i == 0:
                continue
            # リスト入れ替え
            self.files[i-1], self.files[i] = self.files[i], self.files[i-1]
            txt = self.lst.get(i)
            self.lst.delete(i)
            self.lst.insert(i-1, txt)
        # 再選択
        self.lst.selection_clear(0, tk.END)
        for i in [max(0, s-1) for s in sel]:
            self.lst.selection_set(i)

    def on_move_down(self):
        # 選択行を一つ下へ
        sel = list(self.lst.curselection())
        if not sel:
            return
        for i in reversed(sel):
            if i >= self.lst.size() - 1:
                continue
            self.files[i+1], self.files[i] = self.files[i], self.files[i+1]
            txt = self.lst.get(i)
            self.lst.delete(i)
            self.lst.insert(i+1, txt)
        self.lst.selection_clear(0, tk.END)
        for i in [min(self.lst.size()-1, s+1) for s in sel]:
            self.lst.selection_set(i)

    def on_delete_selected(self):
        sel = sorted(self.lst.curselection(), reverse=True)
        for i in sel:
            del self.files[i]
            self.lst.delete(i)

    def on_clear(self):
        self.files.clear()
        self.lst.delete(0, tk.END)

    def on_browse_out(self):
        p = filedialog.asksaveasfilename(
            title="保存先を指定",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if p:
            self.out_var.set(p)

    def _set_default_outname(self):
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        self.out_var.set(str(Path.cwd() / f"merged_by_sheet_{stamp}.xlsx"))

    def on_run(self):
        if not self.files:
            messagebox.showwarning("警告", "結合対象ファイルを追加してください。")
            return

        # ホワイトリスト（カンマ区切り）→ リスト化
        wl_raw = self.ent_whitelist.get().strip()
        sheet_whitelist = [s.strip() for s in wl_raw.split(",")] if wl_raw else []

        # ヘッダー行
        try:
            header_row = int(self.ent_header.get().strip() or "0")
            if header_row < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("エラー", "ヘッダー行番号は0以上の整数で指定してください。")
            return

        out = self.out_var.get().strip()
        if not out:
            messagebox.showwarning("警告", "保存先を指定してください。")
            return
        out_path = Path(out)

        # 実行
        try:
            self.btn_run.configure(state="disabled")
            merge_by_sheet(self.files, sheet_whitelist, header_row, out_path)
            messagebox.showinfo("完了", f"結合が完了しました。\n出力: {out_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"結合に失敗しました。\n{e}")
        finally:
            self.btn_run.configure(state="normal")


# -------------------------------
# エントリーポイント
# -------------------------------
def main():
    # tkinterdnd2 が使えるなら DnD版のTk を、なければ通常Tk
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    ExcelMergerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
