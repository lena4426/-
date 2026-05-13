import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
import os
import database as db
import export_excel as xls
import print_report as pr


# ─────────────────────────────────────────────
#  ЦВЕТА
# ─────────────────────────────────────────────
BG        = "#1e1e2e"
PANEL     = "#2a2a3e"
ACCENT    = "#5b4fd4"   # тёмно-фиолетовый — белый текст читается
ACCENT2   = "#7c6af7"
SUCCESS   = "#16a34a"   # тёмно-зелёный
DANGER    = "#b91c1c"   # тёмно-красный
WARNING   = "#b45309"   # тёмно-янтарный
BTN_GREY  = "#3d3d5c"   # для нейтральных кнопок
TEXT      = "#f1f5f9"
TEXT2     = "#94a3b8"
BORDER    = "#3d3d5c"
TREE_ODD  = "#252538"
TREE_EVEN = "#2a2a3e"
ENTRY_BG  = "#16162a"

FONT_H1   = ("Segoe UI", 18, "bold")
FONT_H2   = ("Segoe UI", 12, "bold")
FONT_BODY = ("Segoe UI", 10)
FONT_MONO = ("Consolas", 10)


def _hover_color(color):
    import colorsys
    c = color.lstrip("#")
    r, g, b = [int(c[i:i+2], 16)/255 for i in (0, 2, 4)]
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    v = min(1.0, v + 0.18)
    r2, g2, b2 = colorsys.hsv_to_rgb(h, s, v)
    return "#{:02x}{:02x}{:02x}".format(int(r2*255), int(g2*255), int(b2*255))


def styled_button(parent, text, command, color=ACCENT, fg="#ffffff", **kwargs):
    """
    Используем tk.Label вместо tk.Button — macOS не перекрашивает Label,
    поэтому цвет фона и текста всегда наш.
    """
    hover = _hover_color(color)
    kwargs.pop("activebackground", None)
    kwargs.pop("activeforeground", None)

    btn = tk.Label(
        parent, text=text,
        bg=color, fg="#ffffff",
        font=("Helvetica", 11, "bold"),
        padx=14, pady=7,
        cursor="hand2",
        **kwargs
    )
    btn.bind("<Enter>",          lambda e: btn.config(bg=hover))
    btn.bind("<Leave>",          lambda e: btn.config(bg=color))
    btn.bind("<ButtonPress-1>",  lambda e: btn.config(bg=hover))
    btn.bind("<ButtonRelease-1>", lambda e: (btn.config(bg=color), command()))
    return btn


def styled_entry(parent, textvariable=None, width=30):
    e = tk.Entry(
        parent, textvariable=textvariable, width=width,
        bg=ENTRY_BG, fg=TEXT, insertbackground=TEXT,
        relief="flat", font=FONT_BODY,
        highlightthickness=1, highlightcolor=ACCENT,
        highlightbackground=BORDER
    )
    return e


def styled_combo(parent, textvariable, values, width=28):
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Dark.TCombobox",
        fieldbackground=ENTRY_BG, background=PANEL,
        foreground=TEXT, arrowcolor=TEXT2,
        bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER
    )
    cb = ttk.Combobox(parent, textvariable=textvariable, values=values,
                      width=width, style="Dark.TCombobox", font=FONT_BODY)
    cb.option_add('*TCombobox*Listbox.background', ENTRY_BG)
    cb.option_add('*TCombobox*Listbox.foreground', TEXT)
    cb.option_add('*TCombobox*Listbox.selectBackground', ACCENT)
    return cb


def make_tree(parent, columns, heights=16):
    style = ttk.Style()
    style.configure("Warehouse.Treeview",
        background=TREE_EVEN, foreground=TEXT,
        fieldbackground=TREE_EVEN,
        rowheight=28, font=FONT_BODY,
        borderwidth=0
    )
    style.configure("Warehouse.Treeview.Heading",
        background=PANEL, foreground=ACCENT2,
        font=FONT_H2, relief="flat"
    )
    style.map("Warehouse.Treeview",
        background=[("selected", ACCENT)],
        foreground=[("selected", TEXT)]
    )
    tree = ttk.Treeview(parent, columns=columns, show="headings",
                        style="Warehouse.Treeview", height=heights)
    tree.tag_configure("odd", background=TREE_ODD)
    tree.tag_configure("even", background=TREE_EVEN)
    return tree


def add_scrollbar(parent, tree):
    sb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    return sb


def refresh_tree(tree, rows, columns):
    tree.delete(*tree.get_children())
    for i, row in enumerate(rows):
        tag = "odd" if i % 2 else "even"
        vals = [row.get(c, "") for c in columns]
        tree.insert("", "end", values=vals, tags=(tag,))


# ═══════════════════════════════════════════════════════════════
#  ДИАЛОГ: СПРАВОЧНИК (общий)
# ═══════════════════════════════════════════════════════════════

class DirectoryDialog(tk.Toplevel):
    def __init__(self, parent, title, table, fields, labels):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=BG)
        self.table = table
        self.fields = fields
        self.labels = labels
        self.resizable(False, False)
        self._selected_id = None
        self._build()
        self._load()

    def _build(self):
        top = tk.Frame(self, bg=BG, padx=16, pady=12)
        top.pack(fill="x")
        tk.Label(top, text=self.title(), font=FONT_H1, bg=BG, fg=TEXT).pack(anchor="w")

        tree_frame = tk.Frame(self, bg=BG, padx=16)
        tree_frame.pack(fill="both", expand=True)
        cols = self.fields
        self.tree = make_tree(tree_frame, cols, heights=12)
        for f, l in zip(self.fields, self.labels):
            self.tree.heading(f, text=l)
            self.tree.column(f, width=200)
        sb = add_scrollbar(tree_frame, self.tree)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        form = tk.Frame(self, bg=PANEL, padx=16, pady=12)
        form.pack(fill="x", padx=16, pady=(0, 8))
        self._vars = {}
        for i, (f, l) in enumerate(zip(self.fields, self.labels)):
            tk.Label(form, text=l + ":", bg=PANEL, fg=TEXT2, font=FONT_BODY).grid(row=i, column=0, sticky="w", pady=3)
            v = tk.StringVar()
            self._vars[f] = v
            styled_entry(form, textvariable=v, width=35).grid(row=i, column=1, padx=(8, 0), pady=3)

        btns = tk.Frame(self, bg=BG, padx=16, pady=8)
        btns.pack(fill="x")
        styled_button(btns, "➕ Добавить", self._add, color=SUCCESS, fg="#ffffff").pack(side="left", padx=4)
        styled_button(btns, "✏️ Сохранить", self._update, color=WARNING, fg="#ffffff").pack(side="left", padx=4)
        styled_button(btns, "🗑 Удалить", self._delete, color=DANGER).pack(side="left", padx=4)

    def _load(self):
        rows = db.get_all(self.table)
        refresh_tree(self.tree, rows, self.fields)

    def _on_select(self, _):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0])["values"]
        for f, v in zip(self.fields, vals):
            self._vars[f].set(v)
        # find id
        rows = db.get_all(self.table)
        for r in rows:
            if all(str(r[f]) == str(self._vars[f].get()) for f in self.fields):
                self._selected_id = r["id"]
                break

    def _add(self):
        kwargs = {f: self._vars[f].get().strip() for f in self.fields}
        if not all(kwargs.values()):
            messagebox.showwarning("Внимание", "Заполните все поля.", parent=self)
            return
        db.add_record(self.table, **kwargs)
        self._load()
        for v in self._vars.values():
            v.set("")

    def _update(self):
        if not self._selected_id:
            messagebox.showwarning("Внимание", "Выберите запись.", parent=self)
            return
        kwargs = {f: self._vars[f].get().strip() for f in self.fields}
        db.update_record(self.table, self._selected_id, **kwargs)
        self._load()

    def _delete(self):
        if not self._selected_id:
            messagebox.showwarning("Внимание", "Выберите запись.", parent=self)
            return
        if messagebox.askyesno("Удаление", "Удалить запись?", parent=self):
            db.delete_record(self.table, self._selected_id)
            self._selected_id = None
            self._load()


# ═══════════════════════════════════════════════════════════════
#  ДИАЛОГ: ДОБАВИТЬ / РЕДАКТИРОВАТЬ документ
# ═══════════════════════════════════════════════════════════════

class DocDialog(tk.Toplevel):
    """Базовый диалог для добавления/редактирования документа."""
    def __init__(self, parent, title, save_callback, initial=None):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=BG)
        self.resizable(False, False)
        self.save_callback = save_callback
        self.initial = initial or {}
        self._build()
        if initial:
            self._fill(initial)

    def _field(self, frame, label, row, widget_factory):
        tk.Label(frame, text=label, bg=PANEL, fg=TEXT2, font=FONT_BODY).grid(
            row=row, column=0, sticky="w", pady=5, padx=(0, 12))
        w = widget_factory()
        w.grid(row=row, column=1, sticky="ew", pady=5)
        return w

    def _build(self):
        raise NotImplementedError

    def _fill(self, data):
        pass

    def _save(self):
        raise NotImplementedError

    def _make_form(self, title_text):
        tk.Label(self, text=title_text, font=FONT_H1, bg=BG, fg=TEXT, padx=16, pady=12).pack(anchor="w")
        frame = tk.Frame(self, bg=PANEL, padx=20, pady=16)
        frame.pack(fill="both", padx=16, pady=(0, 8))
        frame.columnconfigure(1, weight=1)
        return frame

    def _make_buttons(self):
        btns = tk.Frame(self, bg=BG, padx=16, pady=10)
        btns.pack(fill="x")
        styled_button(btns, "💾 Сохранить", self._save, color=SUCCESS, fg="#ffffff").pack(side="left", padx=4)
        styled_button(btns, "✖ Отмена", self.destroy, color=DANGER).pack(side="left", padx=4)


# ═══════════════════════════════════════════════════════════════
#  ДИАЛОГ ЗАКУПКИ
# ═══════════════════════════════════════════════════════════════

class PurchaseDialog(DocDialog):
    def _build(self):
        frame = self._make_form("Закупка")
        self.orgs = db.get_all("organizations")
        self.noms = db.get_all("nomenclature")
        self.whs  = db.get_all("warehouses")

        self.date_var = tk.StringVar(value=str(date.today()))
        self.sup_var  = tk.StringVar()
        self.nom_var  = tk.StringVar()
        self.wh_var   = tk.StringVar()
        self.qty_var  = tk.StringVar()

        self._field(frame, "Дата:", 0, lambda: styled_entry(frame, self.date_var))
        self._field(frame, "Поставщик:", 1, lambda: styled_combo(frame, self.sup_var, [o["name"] for o in self.orgs]))
        self._field(frame, "Номенклатура:", 2, lambda: styled_combo(frame, self.nom_var, [n["name"] for n in self.noms]))
        self._field(frame, "Склад:", 3, lambda: styled_combo(frame, self.wh_var, [w["name"] for w in self.whs]))
        self._field(frame, "Количество:", 4, lambda: styled_entry(frame, self.qty_var, width=15))
        self._make_buttons()

    def _fill(self, d):
        self.date_var.set(d.get("date", ""))
        self.sup_var.set(d.get("supplier", ""))
        self.nom_var.set(d.get("nomenclature", ""))
        self.wh_var.set(d.get("warehouse", ""))
        self.qty_var.set(d.get("quantity", ""))

    def _save(self):
        try:
            sup_id = next(o["id"] for o in self.orgs if o["name"] == self.sup_var.get())
            nom_id = next(n["id"] for n in self.noms if n["name"] == self.nom_var.get())
            wh_id  = next(w["id"] for w in self.whs  if w["name"] == self.wh_var.get())
            qty    = float(self.qty_var.get())
        except (StopIteration, ValueError):
            messagebox.showerror("Ошибка", "Заполните все поля корректно.", parent=self)
            return
        self.save_callback(self.date_var.get(), sup_id, nom_id, wh_id, qty)
        self.destroy()


# ═══════════════════════════════════════════════════════════════
#  ДИАЛОГ ПРОДАЖИ
# ═══════════════════════════════════════════════════════════════

class SaleDialog(DocDialog):
    def _build(self):
        frame = self._make_form("Продажа")
        self.orgs = db.get_all("organizations")
        self.noms = db.get_all("nomenclature")
        self.whs  = db.get_all("warehouses")

        self.date_var = tk.StringVar(value=str(date.today()))
        self.cus_var  = tk.StringVar()
        self.nom_var  = tk.StringVar()
        self.wh_var   = tk.StringVar()
        self.qty_var  = tk.StringVar()

        self._field(frame, "Дата:", 0, lambda: styled_entry(frame, self.date_var))
        self._field(frame, "Покупатель:", 1, lambda: styled_combo(frame, self.cus_var, [o["name"] for o in self.orgs]))
        self._field(frame, "Номенклатура:", 2, lambda: styled_combo(frame, self.nom_var, [n["name"] for n in self.noms]))
        self._field(frame, "Склад:", 3, lambda: styled_combo(frame, self.wh_var, [w["name"] for w in self.whs]))
        self._field(frame, "Количество:", 4, lambda: styled_entry(frame, self.qty_var, width=15))
        self._make_buttons()

    def _fill(self, d):
        self.date_var.set(d.get("date", ""))
        self.cus_var.set(d.get("customer", ""))
        self.nom_var.set(d.get("nomenclature", ""))
        self.wh_var.set(d.get("warehouse", ""))
        self.qty_var.set(d.get("quantity", ""))

    def _save(self):
        try:
            cus_id = next(o["id"] for o in self.orgs if o["name"] == self.cus_var.get())
            nom_id = next(n["id"] for n in self.noms if n["name"] == self.nom_var.get())
            wh_id  = next(w["id"] for w in self.whs  if w["name"] == self.wh_var.get())
            qty    = float(self.qty_var.get())
        except (StopIteration, ValueError):
            messagebox.showerror("Ошибка", "Заполните все поля корректно.", parent=self)
            return
        self.save_callback(self.date_var.get(), cus_id, nom_id, wh_id, qty)
        self.destroy()


# ═══════════════════════════════════════════════════════════════
#  ДИАЛОГ ПЕРЕМЕЩЕНИЯ
# ═══════════════════════════════════════════════════════════════

class MovementDialog(DocDialog):
    def _build(self):
        frame = self._make_form("Перемещение")
        self.noms = db.get_all("nomenclature")
        self.whs  = db.get_all("warehouses")
        wh_names  = [w["name"] for w in self.whs]

        self.date_var  = tk.StringVar(value=str(date.today()))
        self.from_var  = tk.StringVar()
        self.to_var    = tk.StringVar()
        self.nom_var   = tk.StringVar()
        self.qty_var   = tk.StringVar()

        self._field(frame, "Дата:", 0, lambda: styled_entry(frame, self.date_var))
        self._field(frame, "Со склада:", 1, lambda: styled_combo(frame, self.from_var, wh_names))
        self._field(frame, "На склад:", 2, lambda: styled_combo(frame, self.to_var, wh_names))
        self._field(frame, "Номенклатура:", 3, lambda: styled_combo(frame, self.nom_var, [n["name"] for n in self.noms]))
        self._field(frame, "Количество:", 4, lambda: styled_entry(frame, self.qty_var, width=15))
        self._make_buttons()

    def _fill(self, d):
        self.date_var.set(d.get("date", ""))
        self.from_var.set(d.get("from_warehouse", ""))
        self.to_var.set(d.get("to_warehouse", ""))
        self.nom_var.set(d.get("nomenclature", ""))
        self.qty_var.set(d.get("quantity", ""))

    def _save(self):
        try:
            from_id = next(w["id"] for w in self.whs if w["name"] == self.from_var.get())
            to_id   = next(w["id"] for w in self.whs if w["name"] == self.to_var.get())
            nom_id  = next(n["id"] for n in self.noms if n["name"] == self.nom_var.get())
            qty     = float(self.qty_var.get())
        except (StopIteration, ValueError):
            messagebox.showerror("Ошибка", "Заполните все поля корректно.", parent=self)
            return
        if from_id == to_id:
            messagebox.showerror("Ошибка", "Склады должны быть разными.", parent=self)
            return
        self.save_callback(self.date_var.get(), from_id, to_id, nom_id, qty)
        self.destroy()


# ═══════════════════════════════════════════════════════════════
#  ПАНЕЛЬ: СПИСОК ДОКУМЕНТОВ (общий шаблон)
# ═══════════════════════════════════════════════════════════════

class DocPanel(tk.Frame):
    COLUMNS = []
    HEADERS = []

    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self._all_rows = []   # все строки без фильтра поиска
        self._build()
        self.refresh()

    def _build(self):
        # ── Заголовок + кнопки CRUD ──────────────────────────────
        hdr = tk.Frame(self, bg=BG, padx=20, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text=self.TITLE, font=FONT_H1, bg=BG, fg=TEXT).pack(side="left")

        btns = tk.Frame(hdr, bg=BG)
        btns.pack(side="right")
        styled_button(btns, "➕ Добавить",  self._add).pack(side="left", padx=3)
        styled_button(btns, "✏️ Изменить",  self._edit,    color=WARNING).pack(side="left", padx=3)
        styled_button(btns, "🗑 Удалить",   self._delete,  color=DANGER).pack(side="left", padx=3)
        styled_button(btns, "🔄 Обновить",  self.refresh,  color=BTN_GREY).pack(side="left", padx=3)

        # ── Строка поиска + экспорт + печать ─────────────────────
        toolbar = tk.Frame(self, bg=PANEL, padx=16, pady=8)
        toolbar.pack(fill="x", padx=20, pady=(0, 6))

        tk.Label(toolbar, text="🔍", bg=PANEL, fg=TEXT2,
                 font=("Helvetica", 13)).pack(side="left")
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_search())
        search_entry = styled_entry(toolbar, textvariable=self._search_var, width=32)
        search_entry.pack(side="left", padx=(4, 16))
        tk.Label(toolbar, text="Поиск по любому полю",
                 bg=PANEL, fg=TEXT2, font=("Helvetica", 9)).pack(side="left")

        # Кнопки экспорт / печать справа
        styled_button(toolbar, "📊 Excel",  self._export, color="#166534").pack(side="right", padx=3)
        styled_button(toolbar, "🖨 Печать", self._print,  color="#1e3a5f").pack(side="right", padx=3)

        # ── Таблица ───────────────────────────────────────────────
        tf = tk.Frame(self, bg=BG, padx=20)
        tf.pack(fill="both", expand=True, pady=(0, 12))
        self.tree = make_tree(tf, self.COLUMNS, heights=20)
        for col, hdr_txt in zip(self.COLUMNS, self.HEADERS):
            self.tree.heading(col, text=hdr_txt,
                              command=lambda c=col: self._sort_by(c))
        self._set_col_widths()
        sb = add_scrollbar(tf, self.tree)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)

        # Статусная строка
        self._status_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._status_var, bg=BG, fg=TEXT2,
                 font=("Helvetica", 9), anchor="w", padx=24).pack(fill="x")

        self._sort_col = None
        self._sort_rev = False

    def _set_col_widths(self):
        for col in self.COLUMNS:
            self.tree.column(col, width=140, minwidth=60)

    def refresh(self):
        self._all_rows = self._fetch()
        self._apply_search()

    def _apply_search(self):
        query = self._search_var.get().strip().lower() if hasattr(self, "_search_var") else ""
        if query:
            rows = [r for r in self._all_rows
                    if any(query in str(v).lower() for v in r.values())]
        else:
            rows = self._all_rows
        refresh_tree(self.tree, rows, self.COLUMNS)
        total = len(self._all_rows)
        shown = len(rows)
        self._status_var.set(
            f"Показано: {shown} из {total}" if query else f"Всего записей: {total}"
        )

    def _sort_by(self, col):
        """Сортировка по клику на заголовок столбца — числа сортируются как числа."""
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col = col
            self._sort_rev = False

        def sort_key(r):
            val = r.get(col, "")
            try:
                return (0, float(val))      # числа идут числово
            except (ValueError, TypeError):
                return (1, str(val).lower()) # строки — лексикографически

        self._all_rows.sort(key=sort_key, reverse=self._sort_rev)
        self._apply_search()
        arrow = " ▲" if not self._sort_rev else " ▼"
        for c, h in zip(self.COLUMNS, self.HEADERS):
            self.tree.heading(c, text=(h + arrow if c == col else h))

    def _fetch(self):
        return []

    def _visible_rows(self):
        """Строки, которые сейчас видны в таблице (с учётом поиска)."""
        query = self._search_var.get().strip().lower() if hasattr(self, "_search_var") else ""
        if query:
            return [r for r in self._all_rows
                    if any(query in str(v).lower() for v in r.values())]
        return self._all_rows

    def _selected_row(self):
        sel = self.tree.selection()
        if not sel:
            return None
        vals = self.tree.item(sel[0])["values"]
        return dict(zip(self.COLUMNS, vals))

    def _export(self):
        """Переопределяется в подклассах."""
        pass

    def _print(self):
        """Переопределяется в подклассах."""
        pass

    def _add(self): pass
    def _edit(self): pass
    def _delete(self): pass


# ═══════════════════════════════════════════════════════════════
#  ПАНЕЛЬ ЗАКУПОК
# ═══════════════════════════════════════════════════════════════

class PurchasesPanel(DocPanel):
    TITLE   = "📦 Закупки"
    COLUMNS = ["id", "date", "supplier", "nomenclature", "unit", "warehouse", "quantity"]
    HEADERS = ["ID", "Дата", "Поставщик", "Номенклатура", "Ед.", "Склад", "Кол-во"]

    def _fetch(self):
        return db.get_purchases()

    def _export(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "закупки.xlsx")
        n = xls.export_purchases(path)
        messagebox.showinfo("Excel сохранён", f"Файл сохранён ({n} строк):\n{path}")

    def _print(self):
        pr.print_purchases(self._visible_rows())

    def _add(self):
        PurchaseDialog(self, "Новая закупка", self._on_add)

    def _on_add(self, *args):
        db.add_purchase(*args)
        self.refresh()

    def _edit(self):
        row = self._selected_row()
        if not row:
            messagebox.showwarning("Внимание", "Выберите запись.")
            return
        def save(date_val, sup_id, nom_id, wh_id, qty):
            db.update_purchase(row["id"], date_val, sup_id, nom_id, wh_id, qty)
            self.refresh()
        PurchaseDialog(self, "Редактировать закупку", save, initial=row)

    def _delete(self):
        row = self._selected_row()
        if not row:
            messagebox.showwarning("Внимание", "Выберите запись.")
            return
        if messagebox.askyesno("Удаление", f"Удалить закупку #{row['id']}?"):
            db.delete_purchase(row["id"])
            self.refresh()


# ═══════════════════════════════════════════════════════════════
#  ПАНЕЛЬ ПРОДАЖ
# ═══════════════════════════════════════════════════════════════

class SalesPanel(DocPanel):
    TITLE   = "🛒 Продажи"
    COLUMNS = ["id", "date", "customer", "nomenclature", "unit", "warehouse", "quantity"]
    HEADERS = ["ID", "Дата", "Покупатель", "Номенклатура", "Ед.", "Склад", "Кол-во"]

    def _fetch(self):
        return db.get_sales()

    def _export(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "продажи.xlsx")
        n = xls.export_sales(path)
        messagebox.showinfo("Excel сохранён", f"Файл сохранён ({n} строк):\n{path}")

    def _print(self):
        pr.print_sales(self._visible_rows())

    def _add(self):
        SaleDialog(self, "Новая продажа", self._on_add)

    def _on_add(self, *args):
        db.add_sale(*args)
        self.refresh()

    def _edit(self):
        row = self._selected_row()
        if not row:
            messagebox.showwarning("Внимание", "Выберите запись.")
            return
        def save(date_val, cus_id, nom_id, wh_id, qty):
            db.update_sale(row["id"], date_val, cus_id, nom_id, wh_id, qty)
            self.refresh()
        SaleDialog(self, "Редактировать продажу", save, initial=row)

    def _delete(self):
        row = self._selected_row()
        if not row:
            messagebox.showwarning("Внимание", "Выберите запись.")
            return
        if messagebox.askyesno("Удаление", f"Удалить продажу #{row['id']}?"):
            db.delete_sale(row["id"])
            self.refresh()


# ═══════════════════════════════════════════════════════════════
#  ПАНЕЛЬ ПЕРЕМЕЩЕНИЙ
# ═══════════════════════════════════════════════════════════════

class MovementsPanel(DocPanel):
    TITLE   = "🔀 Перемещения"
    COLUMNS = ["id", "date", "from_warehouse", "to_warehouse", "nomenclature", "unit", "quantity"]
    HEADERS = ["ID", "Дата", "Со склада", "На склад", "Номенклатура", "Ед.", "Кол-во"]

    def _fetch(self):
        return db.get_movements()

    def _export(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "перемещения.xlsx")
        n = xls.export_movements(path)
        messagebox.showinfo("Excel сохранён", f"Файл сохранён ({n} строк):\n{path}")

    def _print(self):
        pr.print_movements(self._visible_rows())

    def _add(self):
        MovementDialog(self, "Новое перемещение", self._on_add)

    def _on_add(self, *args):
        db.add_movement(*args)
        self.refresh()

    def _edit(self):
        row = self._selected_row()
        if not row:
            messagebox.showwarning("Внимание", "Выберите запись.")
            return
        def save(date_val, from_id, to_id, nom_id, qty):
            db.update_movement(row["id"], date_val, from_id, to_id, nom_id, qty)
            self.refresh()
        MovementDialog(self, "Редактировать перемещение", save, initial=row)

    def _delete(self):
        row = self._selected_row()
        if not row:
            messagebox.showwarning("Внимание", "Выберите запись.")
            return
        if messagebox.askyesno("Удаление", f"Удалить перемещение #{row['id']}?"):
            db.delete_movement(row["id"])
            self.refresh()


# ═══════════════════════════════════════════════════════════════
#  ПАНЕЛЬ ОСТАТКОВ
# ═══════════════════════════════════════════════════════════════

class StockPanel(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self._all_rows = []
        self._build()
        self.refresh()

    def _build(self):
        hdr = tk.Frame(self, bg=BG, padx=20, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📊 Остатки на складах", font=FONT_H1, bg=BG, fg=TEXT).pack(side="left")
        styled_button(hdr, "🔄 Обновить", self.refresh, color=BTN_GREY).pack(side="right", padx=3)
        styled_button(hdr, "🖨 Печать",   self._print,  color="#1e3a5f").pack(side="right", padx=3)
        styled_button(hdr, "📊 Excel",    self._export, color="#166534").pack(side="right", padx=3)

        flt = tk.Frame(self, bg=PANEL, padx=16, pady=10)
        flt.pack(fill="x", padx=20, pady=(0, 6))

        tk.Label(flt, text="Склад:", bg=PANEL, fg=TEXT2, font=FONT_BODY).pack(side="left")
        self.whs = db.get_all("warehouses")
        self.wh_var = tk.StringVar(value="Все")
        styled_combo(flt, self.wh_var, ["Все"] + [w["name"] for w in self.whs], width=20).pack(side="left", padx=(4, 16))

        tk.Label(flt, text="Номенклатура:", bg=PANEL, fg=TEXT2, font=FONT_BODY).pack(side="left")
        self.noms = db.get_all("nomenclature")
        self.nom_var = tk.StringVar(value="Все")
        styled_combo(flt, self.nom_var, ["Все"] + [n["name"] for n in self.noms], width=24).pack(side="left", padx=4)
        styled_button(flt, "Применить", self.refresh, color=ACCENT).pack(side="left", padx=12)

        # Поиск
        tk.Label(flt, text="🔍", bg=PANEL, fg=TEXT2, font=("Helvetica", 13)).pack(side="left", padx=(16, 2))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_search())
        styled_entry(flt, self._search_var, width=20).pack(side="left")

        tf = tk.Frame(self, bg=BG, padx=20)
        tf.pack(fill="both", expand=True, pady=(0, 4))
        cols = ["nomenclature", "unit", "warehouse", "balance"]
        self.tree = make_tree(tf, cols, heights=20)
        for col, h in zip(cols, ["Номенклатура", "Ед.", "Склад", "Остаток"]):
            self.tree.heading(col, text=h)
            self.tree.column(col, width=200)
        sb = add_scrollbar(tf, self.tree)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)

        self._status_var = tk.StringVar()
        tk.Label(self, textvariable=self._status_var, bg=BG, fg=TEXT2,
                 font=("Helvetica", 9), anchor="w", padx=24).pack(fill="x")

    def refresh(self):
        wh_id  = next((w["id"] for w in self.whs  if w["name"] == self.wh_var.get()),  None)
        nom_id = next((n["id"] for n in self.noms if n["name"] == self.nom_var.get()), None)
        self._all_rows = db.get_stock_report(wh_id, nom_id)
        self._apply_search()

    def _apply_search(self):
        q = self._search_var.get().strip().lower() if hasattr(self, "_search_var") else ""
        cols = ["nomenclature", "unit", "warehouse", "balance"]
        rows = [r for r in self._all_rows
                if any(q in str(v).lower() for v in r.values())] if q else self._all_rows
        refresh_tree(self.tree, rows, cols)
        self._status_var.set(f"Показано: {len(rows)} из {len(self._all_rows)}" if q
                             else f"Всего позиций: {len(self._all_rows)}")

    def _visible_rows(self):
        q = self._search_var.get().strip().lower() if hasattr(self, "_search_var") else ""
        return [r for r in self._all_rows if any(q in str(v).lower() for v in r.values())] if q else self._all_rows

    def _export(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "остатки.xlsx")
        n = xls.export_stock(path)
        messagebox.showinfo("Excel сохранён", f"Файл сохранён ({n} строк):\n{path}")

    def _print(self):
        pr.print_stock(self._visible_rows())



# ═══════════════════════════════════════════════════════════════
#  ПАНЕЛЬ ИСТОРИИ ДВИЖЕНИЙ
# ═══════════════════════════════════════════════════════════════

DOC_TYPE_LABELS = {
    "purchase": "Закупка",
    "sale": "Продажа",
    "movement_in": "Перемещение (приход)",
    "movement_out": "Перемещение (расход)",
}

class HistoryPanel(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self._all_rows = []
        self._build()
        self.refresh()

    def _build(self):
        hdr = tk.Frame(self, bg=BG, padx=20, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📋 История движений", font=FONT_H1, bg=BG, fg=TEXT).pack(side="left")
        styled_button(hdr, "🔄 Обновить", self.refresh, color=BTN_GREY).pack(side="right", padx=3)
        styled_button(hdr, "🖨 Печать",   self._print,  color="#1e3a5f").pack(side="right", padx=3)
        styled_button(hdr, "📊 Excel",    self._export, color="#166534").pack(side="right", padx=3)

        flt = tk.Frame(self, bg=PANEL, padx=16, pady=10)
        flt.pack(fill="x", padx=20, pady=(0, 6))

        tk.Label(flt, text="С:", bg=PANEL, fg=TEXT2, font=FONT_BODY).pack(side="left")
        self.date_from = tk.StringVar()
        styled_entry(flt, self.date_from, width=12).pack(side="left", padx=(4, 12))

        tk.Label(flt, text="По:", bg=PANEL, fg=TEXT2, font=FONT_BODY).pack(side="left")
        self.date_to = tk.StringVar()
        styled_entry(flt, self.date_to, width=12).pack(side="left", padx=(4, 16))

        self.whs = db.get_all("warehouses")
        self.wh_var = tk.StringVar(value="Все")
        tk.Label(flt, text="Склад:", bg=PANEL, fg=TEXT2, font=FONT_BODY).pack(side="left")
        styled_combo(flt, self.wh_var, ["Все"] + [w["name"] for w in self.whs], width=18).pack(side="left", padx=(4, 8))
        styled_button(flt, "Применить", self.refresh, color=ACCENT).pack(side="left", padx=8)

        tk.Label(flt, text="🔍", bg=PANEL, fg=TEXT2, font=("Helvetica", 13)).pack(side="left", padx=(12, 2))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_search())
        styled_entry(flt, self._search_var, width=20).pack(side="left")

        tf = tk.Frame(self, bg=BG, padx=20)
        tf.pack(fill="both", expand=True, pady=(0, 4))
        cols = ["date", "document_type", "nomenclature", "unit", "warehouse", "quantity"]
        self.tree = make_tree(tf, cols, heights=20)
        for col, h in zip(cols, ["Дата", "Тип документа", "Номенклатура", "Ед.", "Склад", "Кол-во"]):
            self.tree.heading(col, text=h)
            self.tree.column(col, width=150)
        sb = add_scrollbar(tf, self.tree)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)

        self._status_var = tk.StringVar()
        tk.Label(self, textvariable=self._status_var, bg=BG, fg=TEXT2,
                 font=("Helvetica", 9), anchor="w", padx=24).pack(fill="x")

    def refresh(self):
        wh_id = next((w["id"] for w in self.whs if w["name"] == self.wh_var.get()), None)
        rows = db.get_movement_history(
            date_from=self.date_from.get() or None,
            date_to=self.date_to.get() or None,
            warehouse_id=wh_id
        )
        for r in rows:
            r["document_type"] = DOC_TYPE_LABELS.get(r["document_type"], r["document_type"])
        self._all_rows = rows
        self._apply_search()

    def _apply_search(self):
        q = self._search_var.get().strip().lower() if hasattr(self, "_search_var") else ""
        cols = ["date", "document_type", "nomenclature", "unit", "warehouse", "quantity"]
        rows = [r for r in self._all_rows if any(q in str(v).lower() for v in r.values())] if q else self._all_rows
        refresh_tree(self.tree, rows, cols)
        self._status_var.set(f"Показано: {len(rows)} из {len(self._all_rows)}" if q
                             else f"Всего записей: {len(self._all_rows)}")

    def _visible_rows(self):
        q = self._search_var.get().strip().lower() if hasattr(self, "_search_var") else ""
        return [r for r in self._all_rows if any(q in str(v).lower() for v in r.values())] if q else self._all_rows

    def _export(self):
        from openpyxl import Workbook
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "история_движений.xlsx")
        rows = self._visible_rows()
        wb = Workbook()
        ws = wb.active
        ws.title = "История"
        ws.append(["Дата", "Тип документа", "Номенклатура", "Ед.", "Склад", "Кол-во"])
        for r in rows:
            ws.append([r["date"], r["document_type"], r["nomenclature"],
                       r["unit"], r["warehouse"], r["quantity"]])
        wb.save(path)
        messagebox.showinfo("Excel сохранён", f"Файл сохранён ({len(rows)} строк):\n{path}")

    def _print(self):
        pr.print_history(self._visible_rows())



# ═══════════════════════════════════════════════════════════════
#  ГЛАВНОЕ ОКНО
# ═══════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Складской учёт")
        self.geometry("1200x720")
        self.configure(bg=BG)
        self.minsize(900, 600)
        db.init_db()
        self._build()

    def _build(self):
        # Боковое меню
        sidebar = tk.Frame(self, bg=PANEL, width=210)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # Логотип
        logo_frame = tk.Frame(sidebar, bg=PANEL, pady=24)
        logo_frame.pack(fill="x")
        tk.Label(logo_frame, text="🏭", font=("Segoe UI", 28), bg=PANEL, fg=ACCENT).pack()
        tk.Label(logo_frame, text="Склад", font=("Segoe UI", 14, "bold"), bg=PANEL, fg=TEXT).pack()
        tk.Label(logo_frame, text="Учётная система", font=("Segoe UI", 8), bg=PANEL, fg=TEXT2).pack()

        tk.Frame(sidebar, bg=BORDER, height=1).pack(fill="x", padx=16, pady=8)

        # Навигация
        self._panels = {}
        self._nav_buttons = {}
        self._current = None

        nav_items = [
            ("purchases",  "📦  Закупки"),
            ("sales",      "🛒  Продажи"),
            ("movements",  "🔀  Перемещения"),
            ("stock",      "📊  Остатки"),
            ("history",    "📋  История"),
        ]

        for key, label in nav_items:
            btn = tk.Label(
                sidebar, text=label, anchor="w", padx=20,
                bg=PANEL, fg=TEXT2,
                font=("Helvetica", 11), cursor="hand2",
                pady=8
            )
            btn.pack(fill="x", pady=1)
            btn.bind("<Button-1>", lambda e, k=key: self._show(k))
            btn.bind("<Enter>",    lambda e, b=btn: b.config(bg=BG, fg=TEXT))
            btn.bind("<Leave>",    lambda e, b=btn, k=key: b.config(
                bg=BG if self._current == k else PANEL,
                fg=ACCENT2 if self._current == k else TEXT2
            ))
            self._nav_buttons[key] = btn

        tk.Frame(sidebar, bg=BORDER, height=1).pack(fill="x", padx=16, pady=8)

        # Справочники
        tk.Label(sidebar, text="СПРАВОЧНИКИ", bg=PANEL, fg=TEXT2,
                 font=("Segoe UI", 8, "bold"), pady=4).pack(anchor="w", padx=16)

        dir_items = [
            ("Организации",   "organizations", ["inn","name"], ["ИНН","Наименование"]),
            ("Склады",        "warehouses",    ["name"],       ["Наименование"]),
            ("Номенклатура",  "nomenclature",  ["name","unit"],["Наименование","Ед.изм."]),
        ]
        for title, table, fields, labels in dir_items:
            def _open(e, t=title, tb=table, f=fields, l=labels):
                DirectoryDialog(self, t, tb, f, l)
            lbl = tk.Label(
                sidebar, text=f"  📁 {title}", anchor="w", padx=20,
                bg=PANEL, fg=TEXT2,
                font=("Helvetica", 10), cursor="hand2", pady=6
            )
            lbl.pack(fill="x", pady=1)
            lbl.bind("<Button-1>", _open)
            lbl.bind("<Enter>", lambda e, b=lbl: b.config(bg=BG, fg=TEXT))
            lbl.bind("<Leave>", lambda e, b=lbl: b.config(bg=PANEL, fg=TEXT2))

        # Контент
        self.content = tk.Frame(self, bg=BG)
        self.content.pack(side="left", fill="both", expand=True)

        self._panels["purchases"]  = PurchasesPanel(self.content)
        self._panels["sales"]      = SalesPanel(self.content)
        self._panels["movements"]  = MovementsPanel(self.content)
        self._panels["stock"]      = StockPanel(self.content)
        self._panels["history"]    = HistoryPanel(self.content)

        self._show("purchases")

    def _show(self, key):
        if self._current:
            self._panels[self._current].pack_forget()
            self._nav_buttons[self._current].config(bg=PANEL, fg=TEXT2)
        self._panels[key].pack(fill="both", expand=True)
        self._nav_buttons[key].config(bg=BG, fg=ACCENT2)
        self._current = key


if __name__ == "__main__":
    app = App()
    app.mainloop()
