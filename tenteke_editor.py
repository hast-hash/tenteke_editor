"""
Tenteke Editor  ―  高機能テキストエディタ
  ・タブで複数ファイルを同時開き (中ボタンクリックで閉じる)
  ・ドラッグ＆ドロップでファイルを開く (pip install tkinterdnd2 が必要)
  ・常に画面下まで行番号を表示
  ・正規表現検索・置換  (Ctrl+H)
  ・自動保存 (バックグラウンドスレッド)
  ・マクロ記録・管理 (Python スクリプト)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter import font as tkfont
import os, re, json, time, threading

# ドラッグ＆ドロップ対応ライブラリ (pip install tkinterdnd2)
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ===========================================================
# 定数
# ===========================================================
AUTOSAVE_DIR = os.path.join(os.path.expanduser("~"), ".tenteke_autosave")
MACRO_FILE   = os.path.join(os.path.expanduser("~"), ".tenteke_macros.json")
EDITOR_FONT  = ("Courier New", 11)
LINE_FONT    = ("Courier New", 11)

LIGHT = dict(
    text_bg="#ffffff", text_fg="#003366", insert="black",
    linenum_bg="#f0f0f0", linenum_fg="#999999",
    ruler_bg="#dddddd", ruler_fg="#555555",
)
DARK = dict(
    text_bg="#1e1e1e", text_fg="#ccddff", insert="white",
    linenum_bg="#252526", linenum_fg="#888888",
    ruler_bg="#333333", ruler_fg="#aaaaaa",
)


# ===========================================================
# エディタペイン（1 タブ = 1 ファイル）
# ===========================================================
class EditorPane(tk.Frame):

    def __init__(self, master, app, filepath=None):
        super().__init__(master)
        self.app      = app
        self.filepath = filepath
        self.modified = False
        self._build()
        if filepath:
            self._load_file(filepath)
        else:
            self.after(100, self._fill_line_numbers)

    # -------- UI --------
    def _build(self):
        # ルーラー
        self.ruler = tk.Label(
            self, anchor="w", font=("Courier", 10),
            bg=LIGHT["ruler_bg"], fg=LIGHT["ruler_fg"],
            text="".join(
                str((i // 10) % 10) if i % 10 == 0 else "·"
                for i in range(1, 201)),
        )
        self.ruler.pack(fill="x")

        body = tk.Frame(self)
        body.pack(fill="both", expand=True)

        # 行番号
        self.line_numbers = tk.Text(
            body, width=5, padx=4, takefocus=0, border=0,
            bg=LIGHT["linenum_bg"], fg=LIGHT["linenum_fg"],
            font=LINE_FONT, state="disabled", cursor="arrow",
            selectbackground=LIGHT["linenum_bg"],
        )
        self.line_numbers.pack(side="left", fill="y")

        # 本文
        self.text = tk.Text(
            body, wrap="none", undo=True,
            bg=LIGHT["text_bg"], fg=LIGHT["text_fg"],
            insertbackground=LIGHT["insert"],
            font=EDITOR_FONT,
            selectbackground="#3390ff", selectforeground="white",
        )
        self.text.pack(side="left", fill="both", expand=True)

        vsb = tk.Scrollbar(body, orient="vertical", command=self._yscroll)
        vsb.pack(side="right", fill="y")
        self.text.config(yscrollcommand=vsb.set)

        hsb = tk.Scrollbar(self, orient="horizontal", command=self.text.xview)
        hsb.pack(side="bottom", fill="x")
        self.text.config(xscrollcommand=hsb.set)

        # イベント
        self.text.bind("<KeyRelease>",      self._on_key)
        self.text.bind("<ButtonRelease-1>", self._on_click)
        self.text.bind("<MouseWheel>",      self._on_wheel)
        self.text.bind("<Configure>",       lambda e: self.after_idle(self._fill_line_numbers))
        self.line_numbers.bind("<MouseWheel>",
            lambda e: self.text.event_generate("<MouseWheel>", delta=e.delta))

    # -------- ファイル --------
    def _load_file(self, path):
        try:
            with open(path, encoding="utf-8", errors="ignore") as f:
                content = f.read()
            self.text.delete("1.0", "end")
            self.text.insert("1.0", content)
            self.filepath = path
            self.modified = False
            self.after(50, self._fill_line_numbers)
        except Exception as e:
            messagebox.showerror("エラー", f"開けません:\n{e}")

    def save(self):
        if not self.filepath:
            return self.save_as()
        return self._write(self.filepath)

    def save_as(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("テキスト","*.txt"),("Python","*.py"),("すべて","*.*")])
        if path:
            self.filepath = path
            return self._write(path)
        return False

    def _write(self, path):
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.text.get("1.0", "end-1c"))
            self.modified = False
            self.app._update_tab_title(self)
            return True
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))
            return False

    # -------- 行番号 --------
    def _fill_line_numbers(self, _=None):
        """テキスト行数 と 画面表示可能行数 の大きい方まで常に番号を表示"""
        total   = int(self.text.index("end-1c").split(".")[0])
        visible = self._visible_lines()
        count   = max(total, visible)

        self.line_numbers.config(state="normal")
        self.line_numbers.delete("1.0", "end")
        self.line_numbers.insert("1.0", "\n".join(str(i) for i in range(1, count + 1)))
        self.line_numbers.config(state="disabled")
        self.line_numbers.yview_moveto(self.text.yview()[0])

    def _visible_lines(self):
        h = self.text.winfo_height()
        if h <= 1:
            return 60
        try:
            f  = tkfont.Font(font=self.text.cget("font"))
            lh = f.metrics("linespace")
            return max(1, h // max(lh, 1)) + 3
        except Exception:
            return 60

    # -------- スクロール --------
    def _yscroll(self, *args):
        self.text.yview(*args)
        self.line_numbers.yview(*args)

    def _on_wheel(self, _=None):
        self.after_idle(lambda: self.line_numbers.yview_moveto(self.text.yview()[0]))

    # -------- イベント --------
    def _on_key(self, event=None):
        self._fill_line_numbers()
        self.app._update_statusbar(self)
        if not self.modified:
            self.modified = True
            self.app._update_tab_title(self)
        # IME 変換中は event.char がスペース(' ')を返すことがある。
        # isprintable() はスペースも True を返すため、明示的に除外する。
        char = event.char
        if (self.app._recording and char
                and len(char) == 1
                and char.isprintable()
                and char != ' '):
            self.app._macro_buffer.append(char)

    def _on_click(self, _=None):
        self.app._update_statusbar(self)
        self.after_idle(lambda: self.line_numbers.yview_moveto(self.text.yview()[0]))

    # -------- テーマ --------
    def apply_theme(self, th):
        self.text.config(bg=th["text_bg"], fg=th["text_fg"], insertbackground=th["insert"])
        self.line_numbers.config(
            bg=th["linenum_bg"], fg=th["linenum_fg"],
            selectbackground=th["linenum_bg"])
        self.ruler.config(bg=th["ruler_bg"], fg=th["ruler_fg"])

    # -------- 自動保存 --------
    def autosave(self):
        content = self.text.get("1.0", "end-1c")
        if not content.strip():
            return None
        os.makedirs(AUTOSAVE_DIR, exist_ok=True)
        base = os.path.splitext(os.path.basename(self.filepath))[0] \
               if self.filepath else "untitled"
        path = os.path.join(AUTOSAVE_DIR, f"{base}_{time.strftime('%Y%m%d_%H%M%S')}.bak")
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        # 古いバックアップを 10 件に制限
        baks = sorted(
            [x for x in os.listdir(AUTOSAVE_DIR) if x.startswith(base)],
            reverse=True)
        for old in baks[10:]:
            try: os.remove(os.path.join(AUTOSAVE_DIR, old))
            except Exception: pass
        return path


# ===========================================================
# 検索・置換ダイアログ
# ===========================================================
class SearchReplaceDialog(tk.Toplevel):

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.title("検索と置換")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self._ranges  = []
        self._current = -1
        self._build()
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _build(self):
        p = dict(padx=6, pady=3)
        tk.Label(self, text="検索:").grid(row=0, column=0, sticky="e", **p)
        self.sv = tk.StringVar()
        tk.Entry(self, textvariable=self.sv, width=36).grid(row=0, column=1, columnspan=3, **p)

        tk.Label(self, text="置換:").grid(row=1, column=0, sticky="e", **p)
        self.rv = tk.StringVar()
        tk.Entry(self, textvariable=self.rv, width=36).grid(row=1, column=1, columnspan=3, **p)

        self.re_v   = tk.BooleanVar()
        self.case_v = tk.BooleanVar()
        self.wrap_v = tk.BooleanVar(value=True)
        tk.Checkbutton(self, text="正規表現",     variable=self.re_v  ).grid(row=2, column=0, **p)
        tk.Checkbutton(self, text="大文字小文字", variable=self.case_v).grid(row=2, column=1, **p)
        tk.Checkbutton(self, text="折り返し検索", variable=self.wrap_v).grid(row=2, column=2, **p)

        bf = tk.Frame(self); bf.grid(row=3, column=0, columnspan=4, pady=6)
        for label, cmd in [("次を検索", self.find_next), ("前を検索", self.find_prev),
                            ("置換", self.replace_one), ("すべて置換", self.replace_all),
                            ("全ハイライト", self.highlight_all)]:
            tk.Button(bf, text=label, width=11, command=cmd).pack(side="left", padx=2)

        self.stv = tk.StringVar()
        tk.Label(self, textvariable=self.stv, fg="navy").grid(
            row=4, column=0, columnspan=4, padx=6, pady=3)

        self.sv.trace_add("write", lambda *_: self._clear())

    @property
    def _text(self):
        p = self.app.current_pane()
        return p.text if p else None

    def _pattern(self):
        raw = self.sv.get()
        if not raw or not self._text: return None
        flags = 0 if self.case_v.get() else re.IGNORECASE
        try:
            return re.compile(raw if self.re_v.get() else re.escape(raw), flags)
        except re.error as e:
            messagebox.showerror("正規表現エラー", str(e), parent=self)
            return None

    def _ofs(self, n):
        content = self._text.get("1.0", "end-1c")
        line = content[:n].count("\n") + 1
        col  = n - (content[:n].rfind("\n") + 1)
        return f"{line}.{col}"

    def _all_matches(self):
        pat = self._pattern()
        if not pat or not self._text: return []
        content = self._text.get("1.0", "end-1c")
        return [(self._ofs(m.start()), self._ofs(m.end())) for m in pat.finditer(content)]

    def highlight_all(self):
        if not self._text: return
        self._clear()
        self._ranges = self._all_matches()
        for s, e in self._ranges:
            self._text.tag_add("hl", s, e)
        self._text.tag_config("hl", background="#ffff00", foreground="black")
        self.stv.set(f"{len(self._ranges)} 件見つかりました")

    def _clear(self):
        if self._text:
            self._text.tag_remove("hl",  "1.0", "end")
            self._text.tag_remove("cur", "1.0", "end")
        self._ranges  = []
        self._current = -1

    def find_next(self): self._find(True)
    def find_prev(self): self._find(False)

    def _find(self, forward):
        self._ranges = self._all_matches()
        if not self._ranges: self.stv.set("見つかりません"); return
        cursor = self._text.index("insert")
        if forward:
            for i, (s, _) in enumerate(self._ranges):
                if self._text.compare(s, ">=", cursor):
                    self._current = i; break
            else:
                if self.wrap_v.get(): self._current = 0
                else: self.stv.set("これ以上ありません"); return
        else:
            found = False
            for i in range(len(self._ranges)-1, -1, -1):
                if self._text.compare(self._ranges[i][0], "<", cursor):
                    self._current = i; found = True; break
            if not found:
                if self.wrap_v.get(): self._current = len(self._ranges)-1
                else: self.stv.set("これ以上ありません"); return
        self._jump(self._current)

    def _jump(self, idx):
        self._text.tag_remove("hl",  "1.0", "end")
        self._text.tag_remove("cur", "1.0", "end")
        for i, (s, e) in enumerate(self._ranges):
            self._text.tag_add("cur" if i == idx else "hl", s, e)
        self._text.tag_config("hl",  background="#ffff00", foreground="black")
        self._text.tag_config("cur", background="#ff6600", foreground="white")
        s, _ = self._ranges[idx]
        self._text.mark_set("insert", s)
        self._text.see(s)
        self.stv.set(f"{idx+1} / {len(self._ranges)} 件")

    def replace_one(self):
        self._ranges = self._all_matches()
        if not self._ranges: self.stv.set("見つかりません"); return
        if self._current < 0 or self._current >= len(self._ranges):
            self._find(True); return
        s, e = self._ranges[self._current]
        self._text.delete(s, e)
        self._text.insert(s, self.rv.get())
        self._ranges = self._all_matches()
        self._find(True)

    def replace_all(self):
        matches = self._all_matches()
        if not matches: self.stv.set("見つかりません"); return
        for s, e in reversed(matches):
            self._text.delete(s, e)
            self._text.insert(s, self.rv.get())
        self.stv.set(f"{len(matches)} 件置換しました")
        self._clear()

    def _close(self):
        self._clear(); self.destroy()


# ===========================================================
# マクロ管理ダイアログ
# ===========================================================
class MacroDialog(tk.Toplevel):

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.title("マクロ管理")
        self.geometry("580x440")
        self.attributes("-topmost", True)
        self.macros = self._load()
        self._build()

    def _load(self):
        if os.path.exists(MACRO_FILE):
            try:
                with open(MACRO_FILE, encoding="utf-8") as f: return json.load(f)
            except Exception: pass
        return {}

    def _save_disk(self):
        with open(MACRO_FILE, "w", encoding="utf-8") as f:
            json.dump(self.macros, f, ensure_ascii=False, indent=2)

    def _build(self):
        lf = tk.Frame(self); lf.pack(side="left", fill="both", padx=6, pady=6)
        tk.Label(lf, text="マクロ一覧").pack()
        self.lb = tk.Listbox(lf, width=22)
        self.lb.pack(fill="both", expand=True)
        self.lb.bind("<<ListboxSelect>>", self._select)
        self._refresh()

        ef = tk.Frame(self); ef.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        tk.Label(ef, text="名前:").pack(anchor="w")
        self.nv = tk.StringVar()
        tk.Entry(ef, textvariable=self.nv).pack(fill="x")
        tk.Label(ef, text="スクリプト (Python):  変数: editor / text / re").pack(
            anchor="w", pady=(6, 0))
        self.sc = tk.Text(ef, font=("Courier New", 10), height=13)
        self.sc.pack(fill="both", expand=True)
        self.sc.insert("1.0",
            "# 例: 選択テキストを大文字に変換\n"
            "try:\n"
            "    sel = editor.get('sel.first', 'sel.last')\n"
            "    editor.delete('sel.first', 'sel.last')\n"
            "    editor.insert('insert', sel.upper())\n"
            "except Exception:\n"
            "    pass\n")
        bf = tk.Frame(ef); bf.pack(pady=4)
        for label, cmd in [("保存", self._save), ("実行", self._run),
                            ("削除", self._delete), ("新規", self._new)]:
            tk.Button(bf, text=label, width=8, command=cmd).pack(side="left", padx=2)
        self.sv = tk.StringVar()
        tk.Label(ef, textvariable=self.sv, fg="navy", wraplength=320).pack()

    def _refresh(self):
        self.lb.delete(0, "end")
        for n in self.macros: self.lb.insert("end", n)

    def _select(self, _=None):
        s = self.lb.curselection()
        if not s: return
        n = self.lb.get(s[0])
        self.nv.set(n)
        self.sc.delete("1.0", "end")
        self.sc.insert("1.0", self.macros[n])

    def _save(self):
        n = self.nv.get().strip()
        if not n: self.sv.set("名前を入力してください"); return
        self.macros[n] = self.sc.get("1.0", "end-1c")
        self._save_disk(); self._refresh()
        self.sv.set(f"「{n}」保存済み")
        self.app._load_macros()

    def _run(self):
        p = self.app.current_pane()
        if not p: return
        script = self.sc.get("1.0", "end-1c")
        try:
            exec(script, {"editor": p.text, "text": p.text.get("1.0","end-1c"),  # noqa: S102
                          "re": re, "tk": tk})
            self.sv.set("実行完了")
        except Exception as e:
            self.sv.set(f"エラー: {e}")

    def _delete(self):
        n = self.nv.get().strip()
        if n in self.macros:
            del self.macros[n]; self._save_disk(); self._refresh()
            self.nv.set(""); self.sc.delete("1.0", "end")
            self.sv.set(f"「{n}」削除")
            self.app._load_macros()

    def _new(self):
        self.nv.set(""); self.sc.delete("1.0", "end")
        self.sc.insert("1.0", "# 新しいマクロをここに記述\n")


# ===========================================================
# メインアプリケーション
# ===========================================================
_BaseApp = TkinterDnD.Tk if HAS_DND else tk.Tk

class TentekeEditor(_BaseApp):

    FRAME_COLORS = ["#ccffff", "#ffccff", "#ffffcc", "#ccffcc"]

    def __init__(self):
        super().__init__()
        self.title("Tenteke Editor")
        self.geometry("1280x820")
        self.minsize(800, 500)

        self.dark_mode        = False
        self._frame_idx       = 0
        self._autosave_stop   = threading.Event()
        self.autosave_enabled = tk.BooleanVar(value=True)
        self.autosave_interval= tk.IntVar(value=60)
        self._recording       = False
        self._macro_buffer    = []
        self._search_dlg      = None
        self._macro_dlg       = None

        os.makedirs(AUTOSAVE_DIR, exist_ok=True)

        self._build_ui()
        self._load_macros()
        self._start_autosave()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.new_tab()   # 起動時に空タブを 1 つ

    # ---------- UI 構築 ----------
    def _build_ui(self):
        self.outer = tk.Frame(self, bg=self.FRAME_COLORS[0], bd=8)
        self.outer.pack(fill="both", expand=True)

        self._create_menu()
        self._create_toolbar()

        self.notebook = ttk.Notebook(self.outer)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self.notebook.bind("<Button-2>", self._close_tab_middle)

        self._create_statusbar()

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)

    def _create_menu(self):
        mb = tk.Menu(self)

        fm = tk.Menu(mb, tearoff=0)
        fm.add_command(label="新規タブ\tCtrl+T",     command=self.new_tab)
        fm.add_command(label="開く\tCtrl+O",          command=self.open_file)
        fm.add_command(label="保存\tCtrl+S",          command=self.save_file)
        fm.add_command(label="名前を付けて保存",      command=self.save_file_as)
        fm.add_command(label="タブを閉じる\tCtrl+W", command=self.close_tab)
        fm.add_separator()
        fm.add_command(label="自動保存設定...",       command=self._autosave_settings)
        fm.add_separator()
        fm.add_command(label="終了",                  command=self._on_close)
        mb.add_cascade(label="ファイル", menu=fm)

        em = tk.Menu(mb, tearoff=0)
        em.add_command(label="元に戻す\tCtrl+Z",
            command=lambda: self.current_pane() and self.current_pane().text.edit_undo())
        em.add_command(label="やり直す\tCtrl+Y",
            command=lambda: self.current_pane() and self.current_pane().text.edit_redo())
        em.add_separator()
        for lbl, ev in [("切り取り","<<Cut>>"), ("コピー","<<Copy>>"), ("貼り付け","<<Paste>>")]:
            em.add_command(label=lbl,
                command=lambda e=ev: self.current_pane() and
                                     self.current_pane().text.event_generate(e))
        em.add_separator()
        em.add_command(label="すべて選択\tCtrl+A",
            command=lambda: self.current_pane() and
                            self.current_pane().text.tag_add("sel","1.0","end"))
        mb.add_cascade(label="編集", menu=em)

        sm = tk.Menu(mb, tearoff=0)
        sm.add_command(label="検索・置換\tCtrl+H", command=self.open_search)
        sm.add_command(label="行ジャンプ\tCtrl+G", command=self.goto_line)
        mb.add_cascade(label="検索", menu=sm)

        self.macro_menu = tk.Menu(mb, tearoff=0)
        self.macro_menu.add_command(label="マクロ管理...",      command=self.open_macro_dialog)
        self.macro_menu.add_command(label="記録開始/停止\tF9", command=self.toggle_record)
        self.macro_menu.add_separator()
        mb.add_cascade(label="マクロ", menu=self.macro_menu)

        vm = tk.Menu(mb, tearoff=0)
        vm.add_command(label="ダークモード\tF5",  command=self.toggle_dark)
        vm.add_command(label="外枠色切替",         command=self.toggle_frame)
        vm.add_separator()
        vm.add_command(label="フォントサイズ変更", command=self.change_font)
        mb.add_cascade(label="表示", menu=vm)

        self.config(menu=mb)

        self.bind_all("<Control-t>", lambda e: self.new_tab())
        self.bind_all("<Control-o>", lambda e: self.open_file())
        self.bind_all("<Control-s>", lambda e: self.save_file())
        self.bind_all("<Control-w>", lambda e: self.close_tab())
        self.bind_all("<Control-h>", lambda e: self.open_search())
        self.bind_all("<Control-g>", lambda e: self.goto_line())
        self.bind_all("<F5>",        lambda e: self.toggle_dark())
        self.bind_all("<F9>",        lambda e: self.toggle_record())

    def _create_toolbar(self):
        tb = tk.Frame(self.outer, bd=1, relief="raised")
        tb.pack(fill="x")

        def B(text, cmd):
            b = tk.Button(tb, text=text, command=cmd,
                          relief="flat", padx=6, pady=2, cursor="hand2")
            b.pack(side="left", padx=1, pady=2)
            return b

        def S():
            tk.Frame(tb, width=1, bg="#aaaaaa").pack(side="left", fill="y", padx=5, pady=3)

        B("📄 新規",       self.new_tab)
        B("📂 開く",       self.open_file)
        B("💾 保存",       self.save_file)
        S()
        B("🔍 検索・置換", self.open_search)
        B("⤵ 行ジャンプ", self.goto_line)
        S()
        self.rec_btn = B("⏺ 記録", self.toggle_record)

        dnd_txt = "  ← ファイルをここにドロップ" if HAS_DND else "  (pip install tkinterdnd2 でD&D対応)"
        tk.Label(tb, text=dnd_txt, fg="#999999", font=("", 9)).pack(side="left")

        self.autosave_lamp = tk.Label(
            tb, text="自動保存: ON", fg="#009900", font=("", 9), padx=8)
        self.autosave_lamp.pack(side="right")

    def _create_statusbar(self):
        bar = tk.Frame(self.outer, bd=1, relief="sunken")
        bar.pack(fill="x", side="bottom")

        self._st_cursor   = tk.Label(bar, text="行: 1  列: 1", anchor="w", padx=8)
        self._st_cursor.pack(side="left")
        tk.Frame(bar, width=1, bg="#aaaaaa").pack(side="left", fill="y", pady=2)
        self._st_chars    = tk.Label(bar, text="文字数: 0", anchor="w", padx=8)
        self._st_chars.pack(side="left")
        tk.Frame(bar, width=1, bg="#aaaaaa").pack(side="left", fill="y", pady=2)
        self._st_autosave = tk.Label(bar, text="", fg="#888888", anchor="w", padx=8)
        self._st_autosave.pack(side="left")
        self._st_rec      = tk.Label(bar, text="", fg="red", anchor="e", padx=8)
        self._st_rec.pack(side="right")

    # ---------- タブ管理 ----------
    def new_tab(self, filepath=None):
        pane  = EditorPane(self.notebook, self, filepath)
        title = os.path.basename(filepath) if filepath else "無題"
        self.notebook.add(pane, text=f"  {title}  ")
        self.notebook.select(pane)
        pane.text.focus_set()
        return pane

    def current_pane(self):
        try:
            w = self.notebook.nametowidget(self.notebook.select())
            return w if isinstance(w, EditorPane) else None
        except Exception:
            return None

    def _update_tab_title(self, pane):
        try:
            idx  = self.notebook.index(pane)
            base = os.path.basename(pane.filepath) if pane.filepath else "無題"
            mark = " ●" if pane.modified else ""
            self.notebook.tab(idx, text=f"  {base}{mark}  ")
        except Exception:
            pass

    def close_tab(self, pane=None):
        p = pane or self.current_pane()
        if not p: return
        if p.modified:
            name = os.path.basename(p.filepath) if p.filepath else "無題"
            ans  = messagebox.askyesnocancel(
                "保存確認", f"「{name}」は変更されています。保存しますか？", parent=self)
            if ans is None: return
            if ans: p.save()
        try: self.notebook.forget(p)
        except Exception: pass
        if self.notebook.index("end") == 0:
            self.new_tab()

    def _close_tab_middle(self, event):
        try:
            idx  = self.notebook.index(f"@{event.x},{event.y}")
            pane = self.notebook.nametowidget(self.notebook.tabs()[idx])
            self.close_tab(pane)
        except Exception:
            pass

    def _on_tab_changed(self, _=None):
        p = self.current_pane()
        if p:
            p.text.focus_set()
            self._update_statusbar(p)

    # ---------- ファイル操作 ----------
    def open_file(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("テキスト","*.txt"),("Python","*.py"),("すべて","*.*")])
        for path in paths:
            self._open_path(path)

    def _open_path(self, path):
        path = path.strip().strip("{}")
        if not os.path.isfile(path): return
        # 既に開いているか
        for tab in self.notebook.tabs():
            pane = self.notebook.nametowidget(tab)
            if isinstance(pane, EditorPane) and pane.filepath == path:
                self.notebook.select(pane); return
        # 空の無題タブを再利用
        p = self.current_pane()
        if p and not p.filepath and not p.modified and p.text.get("1.0","end-1c") == "":
            p._load_file(path)
            self._update_tab_title(p)
        else:
            self.new_tab(path)

    def save_file(self):
        p = self.current_pane()
        if p: p.save()

    def save_file_as(self):
        p = self.current_pane()
        if p:
            p.save_as()
            self._update_tab_title(p)

    # ---------- ドラッグ＆ドロップ ----------
    def _on_drop(self, event):
        raw   = event.data
        # "{パス1} {パス2}" 形式と スペースなし形式の両方に対応
        paths = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        for a, b in paths:
            path = (a or b).strip()
            if path:
                self.after(0, self._open_path, path)

    # ---------- ステータスバー ----------
    def _update_statusbar(self, pane=None):
        p = pane or self.current_pane()
        if not p: return
        row, col = p.text.index("insert").split(".")
        self._st_cursor.config(text=f"行: {row}  列: {int(col)+1}")
        self._st_chars.config(text=f"文字数: {len(p.text.get('1.0','end-1c'))}")

    # ---------- 検索・置換 ----------
    def open_search(self):
        if self._search_dlg and self._search_dlg.winfo_exists():
            self._search_dlg.lift()
        else:
            self._search_dlg = SearchReplaceDialog(self, self)

    def goto_line(self):
        p = self.current_pane()
        if not p: return
        total = int(p.text.index("end-1c").split(".")[0])
        line  = simpledialog.askinteger(
            "行ジャンプ", f"行番号 (1〜{total}):",
            parent=self, minvalue=1, maxvalue=total)
        if line:
            p.text.mark_set("insert", f"{line}.0")
            p.text.see(f"{line}.0")
            self._update_statusbar(p)

    # ---------- マクロ ----------
    def open_macro_dialog(self):
        if self._macro_dlg and self._macro_dlg.winfo_exists():
            self._macro_dlg.lift()
        else:
            self._macro_dlg = MacroDialog(self, self)

    def _load_macros(self):
        macros = {}
        if os.path.exists(MACRO_FILE):
            try:
                with open(MACRO_FILE, encoding="utf-8") as f: macros = json.load(f)
            except Exception: pass
        end = self.macro_menu.index("end")
        if end is not None:
            for i in range(int(end), 2, -1):
                try: self.macro_menu.delete(i)
                except Exception: pass
        for name, script in macros.items():
            self.macro_menu.add_command(
                label=name,
                command=lambda s=script: self._exec_macro(s))

    def _exec_macro(self, script):
        p = self.current_pane()
        if not p: return
        try:
            exec(script, {"editor": p.text, "text": p.text.get("1.0","end-1c"),  # noqa: S102
                          "re": re, "tk": tk})
        except Exception as e:
            messagebox.showerror("マクロエラー", str(e))

    def toggle_record(self):
        self._recording = not self._recording
        if self._recording:
            self._macro_buffer = []
            self.rec_btn.config(text="⏹ 記録中...")
            self._st_rec.config(text="● REC")
        else:
            self.rec_btn.config(text="⏺ 記録")
            self._st_rec.config(text="")
            if self._macro_buffer:
                self._save_recorded()

    def _save_recorded(self):
        name = simpledialog.askstring("マクロ保存", "マクロ名を入力:", parent=self)
        if not name: return
        chars  = "".join(self._macro_buffer).replace("\\","\\\\").replace("'","\\'")
        script = f"editor.insert('insert', '{chars}')"
        macros = {}
        if os.path.exists(MACRO_FILE):
            try:
                with open(MACRO_FILE, encoding="utf-8") as f: macros = json.load(f)
            except Exception: pass
        macros[name] = script
        with open(MACRO_FILE, "w", encoding="utf-8") as f:
            json.dump(macros, f, ensure_ascii=False, indent=2)
        self._load_macros()
        messagebox.showinfo("保存完了", f"「{name}」として保存しました")

    # ---------- 自動保存 ----------
    def _autosave_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("自動保存設定"); dlg.resizable(False, False)
        dlg.attributes("-topmost", True)
        tk.Checkbutton(dlg, text="自動保存を有効にする",
                       variable=self.autosave_enabled).grid(
            row=0, column=0, columnspan=2, padx=12, pady=8, sticky="w")
        tk.Label(dlg, text="保存間隔 (秒):").grid(row=1, column=0, padx=12, pady=4, sticky="e")
        tk.Spinbox(dlg, textvariable=self.autosave_interval,
                   from_=10, to=3600, width=8).grid(row=1, column=1, padx=12)
        tk.Button(dlg, text="OK", width=10,
                  command=lambda: (self._restart_autosave(), dlg.destroy())
                  ).grid(row=2, column=0, columnspan=2, pady=8)

    def _start_autosave(self):
        self._autosave_stop.clear()
        threading.Thread(target=self._autosave_loop, daemon=True).start()

    def _restart_autosave(self):
        self._autosave_stop.set()
        enabled = self.autosave_enabled.get()
        self.autosave_lamp.config(
            text=f"自動保存: {'ON' if enabled else 'OFF'}",
            fg="#009900" if enabled else "#888888")
        self._start_autosave()

    def _autosave_loop(self):
        while not self._autosave_stop.wait(self.autosave_interval.get()):
            if self.autosave_enabled.get():
                self.after(0, self._do_autosave)

    def _do_autosave(self):
        saved = 0
        for tab in self.notebook.tabs():
            pane = self.notebook.nametowidget(tab)
            if isinstance(pane, EditorPane):
                try:
                    if pane.autosave(): saved += 1
                except Exception: pass
        if saved:
            self._st_autosave.config(
                text=f"自動保存: {time.strftime('%H:%M:%S')}  ({saved} タブ)")

    # ---------- テーマ ----------
    def toggle_dark(self):
        self.dark_mode = not self.dark_mode
        th = DARK if self.dark_mode else LIGHT
        for tab in self.notebook.tabs():
            pane = self.notebook.nametowidget(tab)
            if isinstance(pane, EditorPane): pane.apply_theme(th)

    def toggle_frame(self):
        self._frame_idx = (self._frame_idx + 1) % len(self.FRAME_COLORS)
        self.outer.config(bg=self.FRAME_COLORS[self._frame_idx])

    def change_font(self):
        size = simpledialog.askinteger(
            "フォントサイズ", "サイズを入力:",
            parent=self, initialvalue=11, minvalue=6, maxvalue=72)
        if size:
            for tab in self.notebook.tabs():
                pane = self.notebook.nametowidget(tab)
                if isinstance(pane, EditorPane):
                    pane.text.config(font=("Courier New", size))
                    pane.line_numbers.config(font=("Courier New", size))

    # ---------- 終了 ----------
    def _on_close(self):
        for tab in list(self.notebook.tabs()):
            pane = self.notebook.nametowidget(tab)
            if isinstance(pane, EditorPane) and pane.modified:
                self.notebook.select(pane)
                name = os.path.basename(pane.filepath) if pane.filepath else "無題"
                ans  = messagebox.askyesnocancel(
                    "保存確認", f"「{name}」は変更されています。保存しますか？", parent=self)
                if ans is None: return
                if ans: pane.save()
        self._autosave_stop.set()
        self.destroy()


# ===========================================================
# 実行
# ===========================================================
if __name__ == "__main__":
    if not HAS_DND:
        print("ヒント: ドラッグ&ドロップを有効にするには  pip install tkinterdnd2  を実行してください")
    app = TentekeEditor()
    app.mainloop()
