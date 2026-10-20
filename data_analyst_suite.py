# ============================================================
#  DATA ANALYST SUITE — VERSION DESIGN PRO (COMPLET + FIX)
#  ✅ Graphes centrés + responsive (plus de décalage)
#  ✅ Tables pro : lignes alternées + en-têtes clean + largeur auto
#  ✅ Heatmap mieux cadrée (labels non coupés)
#  ✅ k modifiable 3..7 (Spinbox + Appliquer) partagé sur toutes les pages
#  ✅ Menu & Modules en cartes
#  ✅ ZOOM sur tous les graphiques (toolbar Matplotlib)
#  ✅ CYBER : affichage du nombre d’anomalies détectées
#  ✅ Design global : plus pro (ombres, marges, titres, cohérence)
# ============================================================

import os
import sys
import numpy as np
import pandas as pd

os.environ.setdefault("OMP_NUM_THREADS", "1")

# ----------------- Matplotlib -----------------
import matplotlib
matplotlib.use("TkAgg")  # IMPORTANT pour Tkinter
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Circle
from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk

# seaborn (optionnel)
try:
    import seaborn as sns
    _HAS_SNS = True
except Exception:
    _HAS_SNS = False

# SciPy (pour enveloppe convex hull)
try:
    from scipy.spatial import ConvexHull
    _HAS_HULL = True
except Exception:
    _HAS_HULL = False

# ----------------- Sklearn / SciPy -----------------
from sklearn.preprocessing import StandardScaler, MinMaxScaler
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans
from sklearn.ensemble import RandomForestClassifier, IsolationForest
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, confusion_matrix, classification_report
from sklearn.neighbors import LocalOutlierFactor
from scipy.stats import chi2_contingency

# ----------------- Tkinter -----------------
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
import tkinter.font as tkfont
from typing import Optional, Tuple



# ============================================================
#  SAFE FONT (Tkinter font fallback)
# ============================================================

def safe_font(name, size, weight="normal", slant="roman"):
    """
    Safe font factory for Tkinter.
    Avoids font-family parsing bugs (e.g., 'Segoe UI' -> 'UI' treated as size).
    Always returns a robust Windows-safe font tuple.
    """
    base = "Calibri"
    try:
        fam = set(tkfont.families())
        if base not in fam:
            base = "Arial"
    except Exception:
        base = "Arial"
    # Tk expects (family, size, style/weight) or named font; simplest stable tuple:
    style = weight if weight in ("normal", "bold") else "normal"
    return (base, int(size), style)



# ============================================================
#  TABLE SETTINGS (Toggle)
# ============================================================

SHOW_ALL_ROWS = False
DEFAULT_MAX_ROWS = 80

def get_table_max_rows(df_len: int) -> int:
    return df_len if SHOW_ALL_ROWS else DEFAULT_MAX_ROWS


# ============================================================
#  CONFIG
# ============================================================

PROF_NAME = "Dr. EL MKHALET MOUNA"
STUDENT_NAME = "Sara Naim"

DATA_FILENAME = "data/DATA36_Trafic_portuaire_100_individus.xlsx"
CURRENT_K = 4  # modifiable entre 3..7

# ============================================================
#  THÈME PRO (Couleurs + Fonts)
# ============================================================

# Base
COLOR_BG        = "#f5f7fb"
COLOR_CARD      = "#ffffff"
COLOR_BORDER    = "#dbe2ef"
COLOR_TEXT      = "#0f172a"
COLOR_MUTED     = "#64748b"
COLOR_LIGHT     = "#eef2ff"

# Accents modules
COLOR_PRIMARY   = "#0b1220"
COLOR_BLUE      = "#2563eb"
COLOR_VIOLET    = "#7c3aed"
COLOR_GREEN     = "#22c55e"
COLOR_ORANGE    = "#f59e0b"
COLOR_RED       = "#ef4444"

# UI
COLOR_BTN_TEXT        = "#ffffff"
COLOR_TABLE_HEAD_BG   = "#e9effa"   # + pro
COLOR_TABLE_ODD       = "#ffffff"
COLOR_TABLE_EVEN      = "#f1f5ff"   # alternance visible
COLOR_SELECT_BG       = "#dbeafe"

FONT_TITLE        = safe_font("Calibri", 22, "bold")
FONT_SUBTITLE     = safe_font("Calibri", 14, "bold")
FONT_CARD_TITLE  = safe_font("Calibri", 13, "bold")
FONT_TEXT         = safe_font("Calibri", 10, "normal")
FONT_SMALL        = safe_font("Calibri", 9, "normal")
FONT_BUTTON       = safe_font("Calibri", 10, "bold")
FONT_MONO         = safe_font("Consolas", 10, "normal")


# ============================================================
#  GLOBAL UI STYLE (Pro polish)
# ============================================================

def apply_global_style(root_):
    """Applique un style global moderne (ttk + Tk) sans changer l'interface."""
    try:
        root_.configure(bg=COLOR_BG)
        # Fonts par défaut (Tk)
        root_.option_add("*Font", "Calibri 10")
        root_.option_add("*Button.BorderWidth", 0)
        root_.option_add("*Button.Relief", "flat")
        root_.option_add("*Label.Background", COLOR_BG)
    except Exception:
        pass

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure("TFrame", background=COLOR_BG)
    style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT)
    style.configure("TSeparator", background=COLOR_BORDER)

    # Spinbox ttk (k)
    style.configure("TSpinbox",
                    padding=6,
                    relief="flat",
                    borderwidth=1,
                    foreground=COLOR_TEXT,
                    fieldbackground="white",
                    background="white")
    style.map("TSpinbox",
              fieldbackground=[("readonly", "white")],
              bordercolor=[("focus", COLOR_BLUE), ("!focus", COLOR_BORDER)])

    # Scrollbar
    style.configure("Vertical.TScrollbar",
                    gripcount=0,
                    background=COLOR_BORDER,
                    darkcolor=COLOR_BORDER,
                    lightcolor=COLOR_BORDER,
                    troughcolor=COLOR_BG,
                    bordercolor=COLOR_BG,
                    arrowcolor=COLOR_MUTED)
    style.configure("Horizontal.TScrollbar",
                    gripcount=0,
                    background=COLOR_BORDER,
                    darkcolor=COLOR_BORDER,
                    lightcolor=COLOR_BORDER,
                    troughcolor=COLOR_BG,
                    bordercolor=COLOR_BG,
                    arrowcolor=COLOR_MUTED)


# ============================================================
#  PATH (py / exe)
# ============================================================

def resource_path(relative_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)

def resolve_data_file(filename: str) -> str:
    candidates = [
        resource_path(filename),
        os.path.join(os.path.dirname(getattr(sys, "executable", "")), filename),
        os.path.join(os.getcwd(), filename),
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return candidates[0]

DEFAULT_FILE_PATH = resolve_data_file(DATA_FILENAME)

# ============================================================
#  ÉTATS
# ============================================================

MAIN = {
    "path": None,
    "df": None,
    "id_col": None,
    "labels": None,
    "X": None,
    "variables": None,
    "scaler": None,
    "X_scaled": None,
    "pca": None,
    "X_pca": None,
    "explained": None,
    "clusters": None,
    "kmeans": None,
    "rf_model": None,
    "rf_report": None,
}

AFC = {
    "cont": None,
    "freq": None,
    "chi2": None,
    "pval": None,
    "labels_ind": None,
    "labels_var": None
}

CYBER = {
    "df": None,
    "labels": None,
    "X": None,
    "variables": None,
    "scaler": None,
    "X_scaled": None,
    "pca": None,
    "X_pca": None,
    "explained": None,
    "iso_pred": None,
    "lof_pred": None,
}

# ============================================================
#  HELPERS UI
# ============================================================

def clear(frame):
    for w in frame.winfo_children():
        w.destroy()

def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def darken_color(hex_color, factor=0.12):
    r, g, b = _hex_to_rgb(hex_color)
    r = max(0, int(r*(1-factor)))
    g = max(0, int(g*(1-factor)))
    b = max(0, int(b*(1-factor)))
    return f"#{r:02x}{g:02x}{b:02x}"

class Card(tk.Frame):
    """Carte uniforme (même forme partout) + hover soft + shadow légère"""
    def __init__(self, parent, bg=COLOR_CARD, border=COLOR_BORDER, hover=COLOR_LIGHT, **kwargs):
        super().__init__(parent, bg=bg, bd=0, highlightthickness=1, highlightbackground=border, **kwargs)
        self.configure(cursor='hand2')
        self._bg = bg
        self._border = border
        self._hover = hover
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _):
        self.configure(bg=self._hover, highlightbackground=darken_color(self._border, 0.10))
        for child in self.winfo_children():
            try:
                if isinstance(child, (tk.Frame, tk.Label, tk.Button)):
                    child.configure(bg=self._hover)
            except Exception:
                pass

    def _on_leave(self, _):
        self.configure(bg=self._bg, highlightbackground=self._border)
        for child in self.winfo_children():
            try:
                if isinstance(child, (tk.Frame, tk.Label, tk.Button)):
                    child.configure(bg=self._bg)
            except Exception:
                pass

def make_button(parent, text, command, bg, fg=COLOR_BTN_TEXT):
    hover = darken_color(bg, 0.12)
    b = tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=fg,
        activebackground=hover, activeforeground=fg,
        bd=0, relief="flat", cursor="hand2",
        font=FONT_BUTTON, padx=16, pady=9,
        highlightthickness=1, highlightbackground=darken_color(bg, 0.30), highlightcolor=bg
    )
    b.bind("<Enter>", lambda _: b.configure(bg=hover))
    b.bind("<Leave>", lambda _: b.configure(bg=bg))
    return b

# ============================================================
#  TABLE / TEXT / FIGURE (style global)
# ============================================================

def _setup_treeview_style():
    style = ttk.Style()
    style.theme_use("default")

    style.configure(
        "Modern.Treeview",
        font=FONT_TEXT,
        rowheight=28,
        background=COLOR_TABLE_ODD,
        fieldbackground=COLOR_CARD,
        foreground=COLOR_TEXT,
        borderwidth=0,
        relief="flat"
    )

    style.configure(
        "Modern.Treeview.Heading",
        font=safe_font("Calibri", 10, "bold"),
        background=COLOR_TABLE_HEAD_BG,
        foreground=COLOR_TEXT,
        relief="flat"
    )

    style.map(
        "Modern.Treeview",
        background=[("selected", COLOR_SELECT_BG)],
        foreground=[("selected", COLOR_TEXT)]
    )


def show_modern_table(frame, df_show: pd.DataFrame, title: str, max_rows=None):
    clear(frame)
    frame.configure(bg=COLOR_CARD)

    # ✅ barre titre
    topbar = tk.Frame(frame, bg=COLOR_CARD)
    topbar.pack(fill="x", pady=(0, 10))

    tk.Label(
        topbar, text=title, font=FONT_SUBTITLE,
        bg=COLOR_CARD, fg=COLOR_TEXT, anchor="w"
    ).pack(side="left", fill="x", expand=True)

    # ✅ Export Excel (toujours disponible)
    def export_table():
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Tous", "*.*")]
        )
        if path:
            try:
                df_show.to_excel(path, index=False)
                messagebox.showinfo("Export", "✅ Tableau exporté en Excel.")
            except Exception as e:
                messagebox.showerror("Export", f"Erreur export : {e}")

    make_button(topbar, "📥 Export Excel", export_table, bg=COLOR_BLUE).pack(side="right", padx=8)

    # ✅ Toggle UNIQUEMENT pour le fichier des 100 individus
    show_toggle = False
    try:
        p = str(MAIN.get("path", "") or "").lower()
        if ("100_individus" in p) or ("data36_trafic_portuaire_100_individus" in p):
            show_toggle = True
    except Exception:
        show_toggle = False

    if show_toggle:
        def _toggle_rows():
            global SHOW_ALL_ROWS
            SHOW_ALL_ROWS = not SHOW_ALL_ROWS
            show_modern_table(frame, df_show, title)

        btn_txt = "Afficher 80 lignes" if SHOW_ALL_ROWS else "Afficher toutes les lignes"

        toggle_btn = tk.Button(
            topbar, text=btn_txt, command=_toggle_rows,
            bg=COLOR_TABLE_HEAD_BG, fg=COLOR_TEXT,
            activebackground=COLOR_SELECT_BG, activeforeground=COLOR_TEXT,
            bd=0, relief="flat", cursor="hand2",
            font=safe_font("Calibri", 9, "bold"),
            padx=10, pady=6
        )
        toggle_btn.pack(side="right", padx=6)

    _setup_treeview_style()

    container = tk.Frame(frame, bg=COLOR_CARD)
    container.pack(fill="both", expand=True)

    tree = ttk.Treeview(container, style="Modern.Treeview", show="headings")
    vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    container.grid_rowconfigure(0, weight=1)
    container.grid_columnconfigure(0, weight=1)

    df_show = df_show.copy()
    df_show.columns = [str(c) for c in df_show.columns]
    cols = list(df_show.columns)
    tree["columns"] = cols

    tree.tag_configure("odd", background=COLOR_TABLE_ODD)
    tree.tag_configure("even", background=COLOR_TABLE_EVEN)

    max_rows = get_table_max_rows(len(df_show)) if max_rows is None else int(max_rows)
    preview = df_show.head(max_rows)

    for c in cols:
        tree.heading(c, text=c, anchor="center")
        max_len = len(str(c))
        for v in preview[c].astype(str).values[:40]:
            max_len = max(max_len, len(v))
        width = min(340, max(110, int(max_len * 7.5)))
        tree.column(c, width=width, anchor="center", minwidth=90)

    for idx, (_, row) in enumerate(preview.iterrows()):
        tag = "even" if idx % 2 == 0 else "odd"
        tree.insert("", "end", values=list(row), tags=(tag,))

    tk.Label(
        frame,
        text=f"Affichage de {min(len(df_show), max_rows)} lignes sur {len(df_show)}",
        font=FONT_SMALL, bg=COLOR_CARD, fg=COLOR_MUTED
    ).pack(side="bottom", pady=(8, 0))


def show_modern_text(frame, title: str, text: str):
    clear(frame)
    frame.configure(bg=COLOR_CARD)

    tk.Label(frame, text=title, font=FONT_SUBTITLE, bg=COLOR_CARD, fg=COLOR_TEXT, anchor="w")\
        .pack(fill="x", pady=(0, 10))

    box = tk.Frame(frame, bg=COLOR_CARD)
    box.pack(fill="both", expand=True)

    txt = tk.Text(
        box, wrap="word", bg=COLOR_CARD, fg=COLOR_TEXT,
        font=FONT_MONO, bd=1, relief="solid",
        padx=10, pady=10, highlightthickness=1, highlightbackground=COLOR_BORDER
    )
    sb = ttk.Scrollbar(box, orient="vertical", command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    txt.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    txt.insert("1.0", text)
    txt.config(state="disabled")


def _finalize_layout(fig: Figure):
    """Ajuste automatiquement la mise en page pour éviter les titres/labels coupés."""
    try:
        # si beaucoup de labels sur X, on augmente le bas
        for ax in fig.get_axes():
            xt = ax.get_xticklabels()
            if len(xt) >= 8:
                for t in xt:
                    t.set_rotation(35)
                    t.set_ha("right")
                fig.subplots_adjust(left=0.10, right=0.98, top=0.90, bottom=0.25)
                return
        fig.tight_layout(pad=1.2)
    except Exception:
        pass


def show_figure(frame, fig: Figure, title: str, dpi=120, add_toolbar=True):
    """✅ Graphique + Toolbar zoom stable sur TOUS les graphes (fix définitif)"""

    clear(frame)
    frame.configure(bg=COLOR_CARD)

    # ✅ titre
    tk.Label(
        frame, text=title,
        font=FONT_SUBTITLE,
        bg=COLOR_CARD, fg=COLOR_TEXT, anchor="w"
    ).pack(fill="x", pady=(0, 10))

    # ✅ conteneur principal
    main_container = tk.Frame(frame, bg=COLOR_CARD)
    main_container.pack(fill="both", expand=True)

    # ✅ toolbar (en haut => stable)
    toolbar_frame = tk.Frame(main_container, bg=COLOR_CARD)
    toolbar_frame.pack(fill="x", side="top")

    # ✅ graphe
    holder = tk.Frame(main_container, bg=COLOR_CARD)
    holder.pack(fill="both", expand=True)

    fig.set_dpi(dpi)
    fig.patch.set_facecolor(COLOR_CARD)

    for ax in fig.get_axes():
        ax.set_facecolor(COLOR_CARD)
        ax.grid(True, alpha=0.18, linestyle="--")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(COLOR_BORDER)
        ax.spines["bottom"].set_color(COLOR_BORDER)
        ax.tick_params(colors=COLOR_TEXT)

    canvas = FigureCanvasTkAgg(fig, master=holder)
    _finalize_layout(fig)
    canvas.draw()
    widget = canvas.get_tk_widget()
    widget.configure(highlightthickness=0, bg=COLOR_CARD)
    widget.pack(fill="both", expand=True)

    # ✅ toolbar zoom
    if add_toolbar:
        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
        toolbar.update()
        toolbar.pack(fill="x")

        toolbar_frame.update_idletasks()
        _finalize_layout(fig)
        canvas.draw_idle()

    def _resize(_event=None):
        w = holder.winfo_width()
        h = holder.winfo_height()
        if w < 80 or h < 80:
            return
        fig.set_size_inches(w / dpi, h / dpi, forward=True)
        try:
            fig.tight_layout(pad=1.1)
        except Exception:
            pass
        canvas.draw_idle()

    holder.bind("<Configure>", _resize)
    holder.after(100, _resize)

# ============================================================
#  Scrollable Frame
# ============================================================

class ScrollableFrame(tk.Frame):
    def __init__(self, parent, bg=COLOR_BG):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.inner = tk.Frame(self.canvas, bg=bg)
        self.window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")

        self.inner.bind("<Configure>", self._on_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_configure(self, _):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window, width=event.width)

    def _on_mousewheel(self, event):
        try:
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        except Exception:
            pass

# ============================================================
#  DATA PREP
# ============================================================

def read_any_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        df_ = pd.read_excel(path)
    elif ext == ".csv":
        df_ = pd.read_csv(path)
    else:
        df_ = pd.read_excel(path)
    df_.columns = df_.columns.astype(str).str.strip()
    drop_cols = [c for c in df_.columns if c.lower().startswith("unnamed")]
    if drop_cols:
        df_.drop(columns=drop_cols, inplace=True, errors="ignore")
    return df_

def clean_numeric_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str)
         .str.replace("\xa0", "", regex=False)
         .str.replace(" ", "", regex=False)
         .str.replace(",", ".", regex=False),
        errors="coerce"
    )

def detect_id_col(df_: pd.DataFrame) -> Optional[str]:
    prefs = ["Navire", "ID", "Id", "id", "Nom", "Name", "Code", "Ref", "Réf"]
    for c in prefs:
        if c in df_.columns:
            return c
    for c in df_.columns:
        if df_[c].dtype == object or str(df_[c].dtype).startswith("string"):
            return c
    return df_.columns[0] if len(df_.columns) else None

def build_X_auto(df_: pd.DataFrame, id_col: Optional[str], min_cols: int = 2) -> Tuple[pd.DataFrame, list]:
    cols = list(df_.columns)
    if id_col in cols:
        cols.remove(id_col)

    Xc = df_[cols].copy()
    for c in Xc.columns:
        Xc[c] = clean_numeric_series(Xc[c])

    keep = [c for c in Xc.columns if not Xc[c].isna().all()]
    Xc = Xc[keep]

    if Xc.shape[1] < min_cols:
        raise ValueError(
            f"Pas assez de colonnes numériques détectées ({Xc.shape[1]}). "
            f"Il faut au moins {min_cols} colonnes numériques."
        )

    med = Xc.median(numeric_only=True)
    Xc = Xc.fillna(med).fillna(0)
    return Xc, list(Xc.columns)

def compute_afc_state_from_main():
    if MAIN["df"] is None:
        return

    X = MAIN["X"]
    labels = MAIN["labels"]

    n = min(8, len(X))
    p = min(8, X.shape[1])
    X8 = X.iloc[:n, :p].copy()
    vars8 = list(X8.columns.astype(str))
    labels8 = list(labels.iloc[:n].astype(str))

    for c in X8.columns:
        try:
            q = pd.qcut(X8[c], 10, labels=False, duplicates="drop")
            q = q.fillna(0).astype(int) + 1
        except Exception:
            q = pd.cut(X8[c], bins=10, labels=False, duplicates="drop")
            q = q.fillna(0).astype(int) + 1
        X8[c] = q

    cont = X8.values.astype(float)
    total = cont.sum()
    freq = cont / total if total != 0 else cont

    chi2, pval, dof, exp = chi2_contingency(cont)

    AFC.update({
        "cont": cont,
        "freq": freq,
        "chi2": chi2,
        "pval": pval,
        "labels_ind": labels8,
        "labels_var": vars8,
    })

def compute_main_state(df_: pd.DataFrame, path: Optional[str] = None, k: int = 4):
    global CURRENT_K

    id_col = detect_id_col(df_)
    if id_col is not None and id_col in df_.columns:
        labels = df_[id_col].astype(str).fillna("")
        fallback = pd.Series([f"Row_{i+1:03d}" for i in range(len(df_))])
        labels = labels.where(labels.str.len() > 0, other=fallback)
    else:
        labels = pd.Series([f"Row_{i+1:03d}" for i in range(len(df_))])

    X, variables = build_X_auto(df_, id_col=id_col, min_cols=2)

    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    pca = PCA()
    X_pca = pca.fit_transform(X_scaled)
    explained = pca.explained_variance_ratio_

    CURRENT_K = int(k)
    CURRENT_K = max(3, min(7, CURRENT_K))

    if X_pca.shape[1] < 2:
        raise ValueError("ACP: moins de 2 composantes (besoin d'au moins 2 variables numériques).")

    km = KMeans(n_clusters=CURRENT_K, random_state=42, n_init=10)
    clusters = km.fit_predict(X_pca[:, :2])

    X_scaled_df = pd.DataFrame(X_scaled, columns=variables)
    y = clusters
    strat = y if len(np.unique(y)) > 1 else None
    Xtr, Xte, ytr, yte = train_test_split(X_scaled_df, y, test_size=0.30, random_state=42, stratify=strat)

    rf = RandomForestClassifier(n_estimators=200, random_state=42)
    rf.fit(Xtr, ytr)
    pred = rf.predict(Xte)

    rep = (
        f"k = {CURRENT_K}\n"
        f"Accuracy : {accuracy_score(yte, pred):.4f}\n\n"
        f"Rapport de classification :\n{classification_report(yte, pred)}\n"
        f"Matrice de confusion :\n{confusion_matrix(yte, pred)}"
    )

    MAIN.update({
        "path": path,
        "df": df_,
        "id_col": id_col,
        "labels": labels,
        "X": X,
        "variables": variables,
        "scaler": scaler,
        "X_scaled": X_scaled,
        "pca": pca,
        "X_pca": X_pca,
        "explained": explained,
        "clusters": clusters,
        "kmeans": km,
        "rf_model": rf,
        "rf_report": rep,
    })

    compute_afc_state_from_main()

def compute_cyber_state(df_in: pd.DataFrame):
    id_col = detect_id_col(df_in)
    if id_col is not None and id_col in df_in.columns:
        labels = df_in[id_col].astype(str).fillna("")
        fallback = pd.Series([f"Row_{i+1:03d}" for i in range(len(df_in))])
        labels = labels.where(labels.str.len() > 0, other=fallback)
    else:
        labels = pd.Series([f"Row_{i+1:03d}" for i in range(len(df_in))])

    Xc, vars_c = build_X_auto(df_in, id_col=id_col, min_cols=2)

    sc = MinMaxScaler()
    Xc_sc = sc.fit_transform(Xc)

    pca = PCA()
    Xc_pca = pca.fit_transform(Xc_sc)
    expl = pca.explained_variance_ratio_

    n = len(Xc)
    iso = IsolationForest(contamination=0.10, random_state=42)
    iso_pred = iso.fit_predict(Xc_sc) if n >= 2 else np.ones(n, dtype=int)

    nn = max(2, min(20, n - 1))
    lof = LocalOutlierFactor(n_neighbors=nn, contamination=0.10)
    lof_pred = lof.fit_predict(Xc_sc) if n >= 3 else np.ones(n, dtype=int)

    CYBER.update({
        "df": df_in,
        "labels": labels,
        "X": Xc,
        "variables": vars_c,
        "scaler": sc,
        "X_scaled": Xc_sc,
        "pca": pca,
        "X_pca": Xc_pca,
        "explained": expl,
        "iso_pred": iso_pred,
        "lof_pred": lof_pred,
    })

# ============================================================
#  IMPORT + STATUS
# ============================================================

def show_status(msg: str):
    if "status_var" in globals():
        status_var.set(msg)

def import_main_default(silent: bool = False):
    if not os.path.exists(DEFAULT_FILE_PATH):
        if not silent:
            messagebox.showerror(
                "Fichier introuvable",
                f"Je ne trouve pas :\n{DEFAULT_FILE_PATH}\n\n"
                "➡️ Mets le fichier Excel à côté du script .py ou de l'exe.\n"
                "➡️ Sinon clique sur 'Choisir fichier...'"
            )
        return

    try:
        df_ = read_any_table(DEFAULT_FILE_PATH)
        compute_main_state(df_, path=DEFAULT_FILE_PATH, k=CURRENT_K)
        show_status(f"✅ DATA chargée: {os.path.basename(DEFAULT_FILE_PATH)} | k={CURRENT_K} | ID={MAIN['id_col']} | vars={len(MAIN['variables'])}")
        if not silent:
            messagebox.showinfo("Import", "✅ Import OK (data par défaut)")
    except Exception as e:
        show_status(f"❌ Erreur import: {e}")
        if not silent:
            messagebox.showerror("Erreur import", str(e))

def import_main_choose():
    path = filedialog.askopenfilename(
        title="Importer une DATA (Excel/CSV)",
        filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("Tous", "*.*")]
    )
    if not path:
        return
    try:
        df_ = read_any_table(path)
        compute_main_state(df_, path=path, k=CURRENT_K)
        show_status(f"✅ DATA chargée: {os.path.basename(path)} | k={CURRENT_K} | ID={MAIN['id_col']} | vars={len(MAIN['variables'])}")
        messagebox.showinfo("Import", f"✅ Import OK\n{os.path.basename(path)}")
    except Exception as e:
        show_status(f"❌ Erreur import: {e}")
        messagebox.showerror("Erreur import", str(e))

def apply_k_from_widget():
    global CURRENT_K
    try:
        k = int(k_var.get())
    except Exception:
        messagebox.showerror("k", "k doit être un entier.")
        return
    k = max(3, min(7, k))
    k_var.set(str(k))
    CURRENT_K = k
    if MAIN["df"] is None:
        show_status(f"ℹ️ k={CURRENT_K} (importe une DATA pour recalculer)")
        return
    try:
        compute_main_state(MAIN["df"], path=MAIN["path"], k=CURRENT_K)
        show_status(f"✅ Recalcul effectué | k={CURRENT_K} | vars={len(MAIN['variables'])}")
        messagebox.showinfo("k", f"✅ k appliqué : {CURRENT_K}")
    except Exception as e:
        show_status(f"❌ Erreur k: {e}")
        messagebox.showerror("Erreur k", str(e))

# ============================================================
#  ACP
# ============================================================

def _require_main():
    if MAIN["df"] is None:
        messagebox.showwarning("Données", "Importe d'abord les données (boutons en haut).")
        return False
    return True

def acp_stats(target):
    if not _require_main(): return
    X = MAIN["X"]
    stats = pd.DataFrame({
        "Moyenne": X.mean(),
        "Écart-type": X.std(),
        "Min": X.min(),
        "Max": X.max()
    }).round(4).reset_index().rename(columns={"index": "Variable"})
    show_modern_table(target, stats, "ACP — Statistiques descriptives")

def acp_cr(target):
    if not _require_main(): return
    X_scaled = MAIN["X_scaled"]
    vars_ = MAIN["variables"]
    df_cr = pd.DataFrame(X_scaled, columns=vars_).round(4)
    show_modern_table(target, df_cr.head(100), "ACP — Matrice centrée-réduite ")

def acp_corr(target):
    """✅ Heatmap mieux cadrée (pas coupée)"""
    if not _require_main(): return
    corr = MAIN["X"].corr()

    fig = Figure()
    ax = fig.add_subplot(111)
    ax.set_title("ACP — Heatmap corrélations (valeurs affichées)", fontsize=14, fontweight="bold", color=COLOR_TEXT)

    if _HAS_SNS:
        sns.heatmap(
            corr, ax=ax, cmap="coolwarm", vmin=-1, vmax=1,
            annot=True, fmt=".2f", annot_kws={"size": 8},
            linewidths=0.5, linecolor=COLOR_BORDER, cbar=True, square=False
        )
        ax.tick_params(axis='x', rotation=35)
        ax.tick_params(axis='y', rotation=0)
        fig.subplots_adjust(left=0.14, right=0.98, top=0.90, bottom=0.22)
    else:
        im = ax.imshow(corr.values, vmin=-1, vmax=1, cmap="coolwarm")
        fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        ax.set_xticks(range(len(corr.columns)))
        ax.set_xticklabels(corr.columns, rotation=35, ha="right", fontsize=9)
        ax.set_yticks(range(len(corr.index)))
        ax.set_yticklabels(corr.index, fontsize=9)
        for i in range(corr.shape[0]):
            for j in range(corr.shape[1]):
                ax.text(j, i, f"{corr.values[i, j]:.2f}", ha="center", va="center", fontsize=8, color=COLOR_TEXT)

    show_figure(target, fig, "ACP — Heatmap matrice de corrélation")

def acp_inerties(target):
    if not _require_main(): return
    explained = MAIN["explained"]
    cum = np.cumsum(explained)

    fig = Figure()
    ax = fig.add_subplot(111)

    x = np.arange(1, len(explained) + 1)
    ax.bar(x, explained * 100, alpha=0.85)
    ax.plot(x, cum * 100, marker="o", linewidth=2.5)

    ax.set_title("ACP — Inerties (barres) + cumul (courbe)", fontsize=14, fontweight="bold", color=COLOR_TEXT)
    ax.set_xlabel("Composantes")
    ax.set_ylabel("Variance (%)")
    ax.set_xticks(x)

    show_figure(target, fig, "ACP — Inerties & variance expliquée")

def acp_plan(target):
    if not _require_main(): return
    Xp = MAIN["X_pca"]
    labels = MAIN["labels"].astype(str).tolist()
    explained = MAIN["explained"]

    fig = Figure()
    ax = fig.add_subplot(111)
    ax.scatter(Xp[:, 0], Xp[:, 1], s=55, alpha=0.9, edgecolors="white", linewidths=0.8)
    ax.axhline(0, color=COLOR_BORDER, linewidth=1)
    ax.axvline(0, color=COLOR_BORDER, linewidth=1)
    ax.set_xlabel(f"PC1 ({explained[0]*100:.1f}%)")
    ax.set_ylabel(f"PC2 ({explained[1]*100:.1f}%)")
    ax.set_title("ACP — Plan factoriel (individus)", fontsize=14, fontweight="bold", color=COLOR_TEXT)

    n = len(labels)
    idx = np.arange(n) if n <= 25 else np.linspace(0, n-1, 25).astype(int)
    for i in idx:
        ax.text(Xp[i, 0], Xp[i, 1], labels[i], fontsize=8, color=COLOR_TEXT)

    show_figure(target, fig, "ACP — Plan factoriel")

def acp_cercle(target):
    if not _require_main(): return
    pca = MAIN["pca"]
    vars_ = MAIN["variables"]
    eig = pca.explained_variance_
    loadings = pca.components_.T * np.sqrt(eig)

    fig = Figure()
    ax = fig.add_subplot(111)
    ax.add_patch(Circle((0, 0), 1, fill=False, color=COLOR_TEXT, linewidth=2))

    for i, v in enumerate(vars_):
        x, y = loadings[i, 0], loadings[i, 1]
        ax.arrow(0, 0, x, y, head_width=0.03, length_includes_head=True, alpha=0.9)
        ax.text(x * 1.12, y * 1.12, str(v), fontsize=9, color=COLOR_TEXT)

    ax.set_xlim(-1.1, 1.1)
    ax.set_ylim(-1.1, 1.1)
    ax.axhline(0, color=COLOR_BORDER, linewidth=1)
    ax.axvline(0, color=COLOR_BORDER, linewidth=1)
    ax.set_aspect("equal", "box")
    ax.set_title("ACP — Cercle de corrélation (PC1/PC2)", fontsize=14, fontweight="bold", color=COLOR_TEXT)

    show_figure(target, fig, "ACP — Cercle de corrélation")

def acp_quality(target):
    if not _require_main(): return
    Xp = MAIN["X_pca"]
    labels = MAIN["labels"].astype(str).tolist()

    denom = np.sum(Xp**2, axis=1, keepdims=True)
    denom = np.where(denom == 0, 1, denom)
    cos2_ind = (Xp**2) / denom
    df_ind = pd.DataFrame({
        "Individu": labels,
        "COS² PC1": cos2_ind[:, 0],
        "COS² PC2": cos2_ind[:, 1],
        "COS² PC1+PC2": cos2_ind[:, 0] + cos2_ind[:, 1]
    }).round(4)

    pca = MAIN["pca"]
    vars_ = MAIN["variables"]
    eig = pca.explained_variance_
    loadings = pca.components_.T * np.sqrt(eig)
    denv = np.sum(loadings**2, axis=1, keepdims=True)
    denv = np.where(denv == 0, 1, denv)
    cos2_var = (loadings[:, :2]**2) / denv
    df_var = pd.DataFrame({
        "Variable": vars_,
        "COS² PC1": cos2_var[:, 0],
        "COS² PC2": cos2_var[:, 1],
        "COS² PC1+PC2": cos2_var[:, 0] + cos2_var[:, 1]
    }).round(4)

    txt = "ACP — Qualité de représentation (COS²)\n\n"
    txt += "1) Individus (top 30 par COS² PC1+PC2)\n"
    txt += df_ind.sort_values("COS² PC1+PC2", ascending=False).head(30).to_string(index=False)
    txt += "\n\n2) Variables\n"
    txt += df_var.sort_values("COS² PC1+PC2", ascending=False).to_string(index=False)
    show_modern_text(target, "ACP — Qualité de représentation (COS²)", txt)

def acp_contrib(target):
    if not _require_main(): return
    Xp = MAIN["X_pca"]
    labels = MAIN["labels"].astype(str).tolist()
    pca = MAIN["pca"]
    vars_ = MAIN["variables"]

    contrib_var = pd.DataFrame({
        "Variable": vars_,
        "Contrib PC1 (%)": (pca.components_[0, :]**2) * 100,
        "Contrib PC2 (%)": (pca.components_[1, :]**2) * 100
    }).round(3)

    n = len(Xp)
    eig = pca.explained_variance_
    contrib_ind = 100 * (Xp[:, :2]**2) / (n * eig[:2])
    df_ind = pd.DataFrame({
        "Individu": labels,
        "Contrib PC1 (%)": contrib_ind[:, 0],
        "Contrib PC2 (%)": contrib_ind[:, 1]
    }).round(3)

    txt = "ACP — Contributions\n\n"
    txt += "1) Variables\n"
    txt += contrib_var.sort_values("Contrib PC1 (%)", ascending=False).to_string(index=False)
    txt += "\n\n2) Individus (top 30 PC1)\n"
    txt += df_ind.sort_values("Contrib PC1 (%)", ascending=False).head(30).to_string(index=False)

    show_modern_text(target, "ACP — Contributions", txt)

# ============================================================
#  IA (Clusters avec enveloppe + couleurs propres)
# ============================================================

def _cluster_palette(k):
    base = ["#2563eb", "#7c3aed", "#22c55e", "#f59e0b", "#ef4444", "#06b6d4", "#a855f7"]
    return base[:k]

def _draw_cluster_hulls(ax, X2, labels_cluster, colors):
    k = len(colors)
    for c in range(k):
        pts = X2[labels_cluster == c]
        if pts.shape[0] < 3:
            continue
        if _HAS_HULL:
            try:
                hull = ConvexHull(pts)
                poly = pts[hull.vertices]
                poly = np.vstack([poly, poly[0]])
                ax.plot(poly[:, 0], poly[:, 1], linewidth=2.2, alpha=0.95, color=colors[c])
            except Exception:
                pass

def ia_clusters(target):
    if not _require_main(): return
    Xp = MAIN["X_pca"][:, :2]
    cl = MAIN["clusters"]
    labels = MAIN["labels"].astype(str).tolist()

    k = int(CURRENT_K)
    colors = _cluster_palette(k)

    fig = Figure()
    ax = fig.add_subplot(111)

    for c in range(k):
        pts = Xp[cl == c]
        ax.scatter(
            pts[:, 0], pts[:, 1],
            s=70, alpha=0.88,
            color=colors[c],
            edgecolors="white", linewidths=1.2,
            label=f"Cluster {c}"
        )

    _draw_cluster_hulls(ax, Xp, cl, colors)

    ax.set_title(f"IA — Clusters KMeans (k={CURRENT_K})", fontsize=14, fontweight="bold", color=COLOR_TEXT)
    ax.set_xlabel("PC1")
    ax.set_ylabel("PC2")
    ax.legend(fontsize=9, frameon=True)

    n = len(labels)
    idx = np.arange(n) if n <= 22 else np.linspace(0, n-1, 22).astype(int)
    for i in idx:
        ax.text(Xp[i, 0], Xp[i, 1], labels[i], fontsize=8, color=COLOR_TEXT, alpha=0.85)

    show_figure(target, fig, "IA — Visualisation des clusters (contours)")

def ia_pourcentage(target):
    if not _require_main(): return
    cl = MAIN["clusters"]
    s = pd.Series(cl).value_counts(normalize=True).sort_index() * 100
    dfp = s.reset_index()
    dfp.columns = ["Cluster", "%"]
    dfp["Cluster"] = dfp["Cluster"].astype(int)
    show_modern_table(target, dfp.round(2), "IA — Répartition des clusters (%)", max_rows=50)

def ia_rf_metrics(target):
    if not _require_main(): return
    show_modern_text(target, "IA — Métriques Random Forest", MAIN["rf_report"])

def ia_predict_new(target):
    if not _require_main(): return
    clear(target)
    target.configure(bg=COLOR_CARD)

    tk.Label(target, text="IA — Prédire un nouvel individu (Random Forest)",
             font=FONT_SUBTITLE, bg=COLOR_CARD, fg=COLOR_TEXT, anchor="w").pack(fill="x", pady=(0, 10))

    X = MAIN["X"]
    vars_ = MAIN["variables"]
    med = X.median(numeric_only=True)

    form = tk.Frame(target, bg=COLOR_CARD)
    form.pack(fill="x", padx=10, pady=10)

    entries = {}
    for v in vars_:
        row = tk.Frame(form, bg=COLOR_CARD)
        row.pack(fill="x", pady=4)

        tk.Label(row, text=str(v), bg=COLOR_CARD, fg=COLOR_TEXT,
                 font=FONT_TEXT, width=28, anchor="w").pack(side="left")

        e = tk.Entry(row, font=FONT_MONO, width=18, bd=1, relief="solid")
        try:
            e.insert(0, str(float(med[v])))
        except Exception:
            e.insert(0, "0")
        e.pack(side="left")
        entries[v] = e

    out = tk.Label(target, text="", bg=COLOR_CARD, fg=COLOR_VIOLET, font=safe_font("Calibri", 12, "bold"))
    out.pack(pady=10)

    def do_predict():
        try:
            vals = [float(entries[v].get().replace(",", ".")) for v in vars_]
        except Exception:
            messagebox.showerror("Erreur", "Merci de saisir uniquement des nombres.")
            return
        x_new = pd.DataFrame([vals], columns=vars_)
        x_sc = MAIN["scaler"].transform(x_new)
        pred = MAIN["rf_model"].predict(x_sc)[0]
        out.config(text=f"➡️ Cluster prédit : {int(pred)} (k={CURRENT_K})")

    make_button(target, "Prédire", do_predict, bg=COLOR_BLUE).pack(pady=8)

# ============================================================
#  AFC
# ============================================================

def _require_afc():
    if MAIN["df"] is None:
        messagebox.showwarning("Données", "Importe une DATA d'abord.")
        return False
    if AFC["cont"] is None:
        messagebox.showwarning("AFC", "AFC n'est pas prête (pas assez de données).")
        return False
    return True

def afc_cont(target):
    if not _require_afc(): return
    cont = AFC["cont"]
    dfc = pd.DataFrame(cont, columns=AFC["labels_var"])
    dfc.insert(0, "Individu", AFC["labels_ind"])
    show_modern_table(target, dfc, "AFC — Tableau de contingence (8×8)")

def afc_freq(target):
    if not _require_afc(): return
    freq = AFC["freq"]
    dff = pd.DataFrame(freq, columns=AFC["labels_var"])
    dff.insert(0, "Individu", AFC["labels_ind"])
    show_modern_table(target, dff.round(4), "AFC — Matrice des fréquences")

def afc_dependance(target):
    if not _require_afc(): return
    cont = AFC["cont"]
    chi2 = AFC["chi2"]
    pval = AFC["pval"]
    n = cont.sum()
    inertia = (chi2 / n) if n != 0 else np.nan
    txt = "AFC — Analyse dépendance (χ²)\n\n"
    txt += f"Somme totale n = {n:.4f}\n"
    txt += f"Chi2 = {chi2:.4f}\n"
    txt += f"p-value = {pval:.6f}\n"
    txt += f"Inertie totale (Chi2/n) = {inertia:.6f}\n\n"
    txt += "Interprétation:\n"
    txt += "- p-value < 0.05 : dépendance significative\n"
    txt += "- p-value ≥ 0.05 : pas de dépendance significative\n"
    show_modern_text(target, "AFC — Dépendance χ² & inertie", txt)

def afc_distance(target):
    if not _require_afc(): return
    cont = AFC["cont"]
    n = cont.sum()
    if n == 0:
        show_modern_text(target, "AFC — Distance χ²", "Contingence vide.")
        return

    P = cont / n
    r = P.sum(axis=1)
    c = P.sum(axis=0)
    r = np.where(r == 0, 1e-12, r)
    c = np.where(c == 0, 1e-12, c)

    row_prof = P / r[:, None]
    d2 = np.sum(((row_prof - c[None, :])**2) / c[None, :], axis=1)
    d = np.sqrt(d2)

    out = pd.DataFrame({"Individu": AFC["labels_ind"], "Distance χ²": d}).round(6)
    show_modern_table(target, out, "AFC — Distance χ² (lignes)")

def afc_plan(target):
    if not _require_afc(): return
    FREQ = AFC["freq"]
    r = FREQ.sum(axis=1)
    c = FREQ.sum(axis=0)
    r = np.where(r == 0, 1e-12, r)
    c = np.where(c == 0, 1e-12, c)

    S = np.diag(1/np.sqrt(r)) @ (FREQ - np.outer(r, c)) @ np.diag(1/np.sqrt(c))
    U, s, Vt = np.linalg.svd(S, full_matrices=False)

    F = np.diag(1/np.sqrt(r)) @ U @ np.diag(s)
    G = np.diag(1/np.sqrt(c)) @ Vt.T @ np.diag(s)

    fig = Figure()
    ax = fig.add_subplot(111)
    ax.scatter(F[:, 0], F[:, 1], s=85, alpha=0.9, edgecolors="white", linewidths=1.0, label="Individus")
    ax.scatter(G[:, 0], G[:, 1], s=95, marker="D", alpha=0.9, edgecolors="white", linewidths=1.0, label="Variables")
    ax.axhline(0, color=COLOR_BORDER, linewidth=1)
    ax.axvline(0, color=COLOR_BORDER, linewidth=1)
    ax.set_title("AFC — Plan factoriel", fontsize=14, fontweight="bold", color=COLOR_TEXT)

    for i, lab in enumerate(AFC["labels_ind"]):
        ax.text(F[i, 0], F[i, 1], str(lab), fontsize=9, color=COLOR_TEXT)
    for j, lab in enumerate(AFC["labels_var"]):
        ax.text(G[j, 0], G[j, 1], str(lab), fontsize=9, color=COLOR_TEXT)

    ax.legend()
    show_figure(target, fig, "AFC — Plan factoriel")

def afc_test(target):
    if not _require_afc(): return
    txt = "AFC — Test χ² & interprétation\n\n"
    txt += f"Chi2 = {AFC['chi2']:.4f}\n"
    txt += f"p-value = {AFC['pval']:.6f}\n\n"
    txt += "Interprétation:\n"
    txt += "- p-value < 0.05 : dépendance significative\n"
    txt += "- p-value ≥ 0.05 : pas de dépendance significative\n"
    show_modern_text(target, "AFC — Test χ²", txt)

# ============================================================
#  CYBER
# ============================================================

def cyber_import(target):
    path = filedialog.askopenfilename(
        title="Importer une DATA CYBER (Excel/CSV)",
        filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("Tous", "*.*")]
    )
    if not path:
        return
    try:
        df_ = read_any_table(path)
        compute_cyber_state(df_)
        show_modern_table(target, CYBER["df"].head(80), f"CYBER — Données importées ({os.path.basename(path)})")
        show_status(f"✅ Données CYBER chargées: {os.path.basename(path)}")
    except Exception as e:
        messagebox.showerror("CYBER import", str(e))

def _require_cyber():
    if CYBER["df"] is None:
        messagebox.showwarning("CYBER", "Importe d'abord des données CYBER (bouton Import).")
        return False
    return True

def cyber_if(target):
    if not _require_cyber(): return
    Xp = CYBER["X_pca"][:, :2]
    labels = CYBER["labels"].astype(str).tolist()
    expl = CYBER["explained"]
    pred = CYBER["iso_pred"]

    fig = Figure()
    ax = fig.add_subplot(111)

    colors = [COLOR_RED if p == -1 else COLOR_BLUE for p in pred]
    ax.scatter(Xp[:, 0], Xp[:, 1], c=colors, s=70, alpha=0.88, edgecolors="white", linewidths=1.1)

    nb = int(np.sum(pred == -1))
    ax.text(0.02, 0.98, f"Anomalies détectées: {nb}",
            transform=ax.transAxes, fontsize=11, fontweight="bold",
            verticalalignment="top", color=COLOR_RED)

    ax.set_title("CYBER — Isolation Forest (anomalies en rouge)", fontsize=14, fontweight="bold", color=COLOR_TEXT)
    ax.set_xlabel(f"PC1 ({expl[0]*100:.1f}%)" if len(expl) > 0 else "PC1")
    ax.set_ylabel(f"PC2 ({expl[1]*100:.1f}%)" if len(expl) > 1 else "PC2")

    idx = np.arange(len(labels)) if len(labels) <= 22 else np.linspace(0, len(labels)-1, 22).astype(int)
    for i in idx:
        ax.text(Xp[i, 0], Xp[i, 1], labels[i], fontsize=8, color=COLOR_TEXT, alpha=0.85)

    show_figure(target, fig, "CYBER — Isolation Forest")

def cyber_lof(target):
    if not _require_cyber(): return
    Xp = CYBER["X_pca"][:, :2]
    labels = CYBER["labels"].astype(str).tolist()
    expl = CYBER["explained"]
    pred = CYBER["lof_pred"]

    fig = Figure()
    ax = fig.add_subplot(111)

    colors = [COLOR_RED if p == -1 else COLOR_BLUE for p in pred]
    ax.scatter(Xp[:, 0], Xp[:, 1], c=colors, s=70, alpha=0.88, edgecolors="white", linewidths=1.1)

    nb = int(np.sum(pred == -1))
    ax.text(0.02, 0.98, f"Anomalies détectées: {nb}",
            transform=ax.transAxes, fontsize=11, fontweight="bold",
            verticalalignment="top", color=COLOR_RED)

    ax.set_title("CYBER — LOF (anomalies en rouge)", fontsize=14, fontweight="bold", color=COLOR_TEXT)
    ax.set_xlabel(f"PC1 ({expl[0]*100:.1f}%)" if len(expl) > 0 else "PC1")
    ax.set_ylabel(f"PC2 ({expl[1]*100:.1f}%)" if len(expl) > 1 else "PC2")

    idx = np.arange(len(labels)) if len(labels) <= 22 else np.linspace(0, len(labels)-1, 22).astype(int)
    for i in idx:
        ax.text(Xp[i, 0], Xp[i, 1], labels[i], fontsize=8, color=COLOR_TEXT, alpha=0.85)

    show_figure(target, fig, "CYBER — Local Outlier Factor")

def cyber_interp(target):
    if not _require_cyber(): return
    labels = CYBER["labels"].astype(str)
    idx_if = np.where(CYBER["iso_pred"] == -1)[0]
    idx_lof = np.where(CYBER["lof_pred"] == -1)[0]

    txt = "CYBER — Interprétation (profils à risque)\n\n"
    txt += "-1 = anomalie | 1 = normal\n\n"
    txt += f"Isolation Forest: {len(idx_if)} anomalies\n"
    txt += (", ".join([labels.iloc[i] for i in idx_if[:20]]) if len(idx_if) else "Aucune") + "\n\n"
    txt += f"LOF: {len(idx_lof)} anomalies\n"
    txt += (", ".join([labels.iloc[i] for i in idx_lof[:20]]) if len(idx_lof) else "Aucune") + "\n\n"
    txt += "Actions recommandées:\n"
    txt += "1) Vérifier les valeurs extrêmes / incohérences\n"
    txt += "2) Contrôler la qualité de saisie (doublons, erreurs)\n"
    txt += "3) Mettre en place un suivi/monitoring si données sensibles\n"

    show_modern_text(target, "CYBER — Interprétation & actions", txt)

# ============================================================
#  PAGES GUI
# ============================================================

frames = {}
root = None
container = None
k_var = None  # sera créé dans init_application()

def make_centered_page(name: str):
    page = tk.Frame(container, bg=COLOR_BG)
    frames[name] = page
    page.grid(row=0, column=0, sticky="nsew")

    center = tk.Frame(page, bg=COLOR_BG)
    center.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.98, relheight=0.96)
    return page, center

def show_frame(name: str):
    if name in frames:
        frames[name].tkraise()

def create_start_page():
    start, c = make_centered_page("START")
    start.configure(bg=COLOR_PRIMARY)

    header = tk.Frame(c, bg=COLOR_PRIMARY)
    header.pack(fill="x", pady=(40, 10))

    tk.Label(header, text="📊", font=safe_font("Calibri", 52, "bold"), bg=COLOR_PRIMARY, fg="white").pack()
    tk.Label(header, text="DATA ANALYST SUITE", font=safe_font("Calibri", 28, "bold"),
             bg=COLOR_PRIMARY, fg="white").pack(pady=(8, 4))
    tk.Label(header, text="ACP • IA • AFC • CYBER", font=safe_font("Calibri", 13, "normal"),
             bg=COLOR_PRIMARY, fg="#c7d2fe").pack()

    card = tk.Frame(c, bg=COLOR_CARD, padx=44, pady=36, bd=0, relief="flat")
    card.place(relx=0.5, rely=0.54, anchor="center")

    tk.Label(card, text="Projet Final", font=safe_font("Calibri", 18, "bold"), bg=COLOR_CARD, fg=COLOR_TEXT).pack(pady=(0, 10))
    tk.Label(card, text="Clique sur Start pour accéder au menu.",
             font=safe_font("Calibri", 11, "normal"), bg=COLOR_CARD, fg=COLOR_MUTED).pack(pady=(0, 18))

    start_btn = make_button(card, "🚀 START", lambda: show_frame("MENU"), bg=COLOR_VIOLET)
    start_btn.config(font=safe_font("Calibri", 14, "bold"), padx=35, pady=12)
    start_btn.pack()

    footer = tk.Frame(c, bg=COLOR_PRIMARY)
    footer.pack(side="bottom", fill="x", pady=(0, 26))

    f = tk.Frame(footer, bg=COLOR_PRIMARY)
    f.pack()

    p = tk.Frame(f, bg=COLOR_CARD, padx=18, pady=12)
    p.pack(side="left", padx=8)
    tk.Label(p, text="👨‍🏫 Prof", font=safe_font("Calibri", 10, "bold"), bg=COLOR_CARD, fg=COLOR_MUTED).pack()
    tk.Label(p, text=PROF_NAME, font=safe_font("Calibri", 11, "normal"), bg=COLOR_CARD, fg=COLOR_TEXT).pack()

    s = tk.Frame(f, bg=COLOR_CARD, padx=18, pady=12)
    s.pack(side="left", padx=8)
    tk.Label(s, text="👨‍🎓 Étudiant", font=safe_font("Calibri", 10, "bold"), bg=COLOR_CARD, fg=COLOR_MUTED).pack()
    tk.Label(s, text=STUDENT_NAME, font=safe_font("Calibri", 11, "normal"), bg=COLOR_CARD, fg=COLOR_TEXT).pack()

def create_menu_page():
    menu, c = make_centered_page("MENU")
    menu.configure(bg=COLOR_BG)

    tk.Label(c, text="MENU PRINCIPAL", font=FONT_TITLE, bg=COLOR_BG, fg=COLOR_TEXT).pack(pady=(10, 5))
    tk.Label(c, text="Sélectionnez un module", font=("Calibri", 12), bg=COLOR_BG, fg=COLOR_MUTED).pack(pady=(0, 18))

    grid = tk.Frame(c, bg=COLOR_BG)
    grid.pack(fill="both", expand=True, padx=40, pady=10)
    grid.grid_columnconfigure(0, weight=1)
    grid.grid_columnconfigure(1, weight=1)
    grid.grid_rowconfigure(0, weight=1)
    grid.grid_rowconfigure(1, weight=1)

    modules = [
        ("☑ ANALYSE ACP", "Visualisation ACP + tables", COLOR_BLUE, "ACP"),
        ("🧠 INTELLIGENCE ARTIFICIELLE", "KMeans + Random Forest", COLOR_VIOLET, "IA"),
        ("📊 ANALYSE AFC", "Contingence + χ² + plan", COLOR_GREEN, "AFC"),
        ("🔐 CYBERSÉCURITÉ", "Isolation Forest + LOF", COLOR_ORANGE, "CYBER"),
    ]

    for i, (title, desc, col, dest) in enumerate(modules):
        r, cc = divmod(i, 2)
        card = Card(grid, bg=COLOR_CARD, border=COLOR_BORDER, hover=COLOR_LIGHT)
        card.grid(row=r, column=cc, sticky="nsew", padx=10, pady=10)

        inner = tk.Frame(card, bg=COLOR_CARD)
        inner.pack(fill="both", expand=True, padx=22, pady=20)

        tk.Label(inner, text=title, font=FONT_SUBTITLE, bg=COLOR_CARD, fg=col, anchor="w")\
            .pack(fill="x")
        tk.Label(inner, text=desc, font=FONT_TEXT, bg=COLOR_CARD, fg=COLOR_MUTED,
                 wraplength=320, justify="left", anchor="w")\
            .pack(fill="x", pady=(8, 14))

        btn = make_button(inner, "Accéder →", lambda d=dest: show_frame(d), bg=col)
        btn.pack(side="right")

    bar = tk.Frame(c, bg=COLOR_PRIMARY)
    bar.pack(side="bottom", fill="x")
    global status_var
    status_var = tk.StringVar(value="ℹ️ Charge une DATA avec Import par défaut ou Choisir fichier...")
    tk.Label(bar, textvariable=status_var, bg=COLOR_PRIMARY, fg="white", font=FONT_SMALL).pack(pady=8)

def create_module_page(name, title, buttons_config, accent_color):
    page, center = make_centered_page(name)
    page.configure(bg=COLOR_BG)

    header = tk.Frame(center, bg=COLOR_BG)
    header.pack(fill="x", padx=20, pady=(10, 0))

    back = tk.Button(header, text="← Retour au menu", command=lambda: show_frame("MENU"),
                     bg=COLOR_BG, fg=accent_color, bd=0, font=safe_font("Calibri", 10, "bold"), cursor="hand2")
    back.pack(side="left")

    tk.Label(header, text=title, font=FONT_TITLE, bg=COLOR_BG, fg=COLOR_TEXT).pack(side="left", padx=20)

    toolbar = tk.Frame(center, bg=COLOR_BG)
    toolbar.pack(fill="x", padx=20, pady=(10, 10))

    ttk.Separator(center, orient='horizontal').pack(fill='x', padx=20, pady=(0, 12))

    make_button(toolbar, "📁 Import par défaut", import_main_default, bg=COLOR_GREEN).pack(side="left", padx=(0, 10))
    make_button(toolbar, "📂 Choisir fichier...", import_main_choose, bg=COLOR_BLUE).pack(side="left")

    tk.Label(toolbar, text="   k (clusters):", bg=COLOR_BG, fg=COLOR_TEXT, font=safe_font("Calibri", 10, "bold"))\
        .pack(side="left", padx=(18, 6))

    spin = ttk.Spinbox(toolbar, from_=3, to=7, textvariable=k_var, width=5, font=FONT_TEXT, justify="center")
    spin.pack(side="left")

    make_button(toolbar, "Appliquer k", apply_k_from_widget, bg=COLOR_PRIMARY).pack(side="left", padx=8)

    wrapper = tk.Frame(center, bg=COLOR_BG)
    wrapper.pack(fill="both", expand=True, padx=20, pady=(0, 15))

    dashboard = tk.Frame(wrapper, bg=COLOR_BG)
    output = tk.Frame(wrapper, bg=COLOR_CARD, bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER)

    dashboard.place(relx=0, rely=0, relwidth=1, relheight=1)
    output.place(relx=0, rely=0, relwidth=1, relheight=1)

    out_head = tk.Frame(output, bg=COLOR_BG)
    out_head.pack(fill="x")

    tk.Button(out_head, text="↩ Retour", command=lambda: dashboard.tkraise(),
              bg=COLOR_BG, fg=accent_color, bd=0,
              font=safe_font("Calibri", 10, "bold"), cursor="hand2").pack(side="left", padx=10, pady=10)

    out_title = tk.Label(out_head, text="RÉSULTATS", bg=COLOR_BG, fg=COLOR_TEXT, font=safe_font("Calibri", 12, "bold"))
    out_title.pack(side="left", padx=10)

    results_content = tk.Frame(output, bg=COLOR_CARD)
    results_content.pack(fill="both", expand=True, padx=12, pady=12)

    def run_action(btn_text, func):
        output.tkraise()
        out_title.config(text=f"RÉSULTATS — {btn_text}")
        func(results_content)

    tk.Label(dashboard, text="Choisissez une analyse", font=FONT_SUBTITLE, bg=COLOR_BG, fg=COLOR_TEXT)\
        .pack(anchor="w", pady=(0, 10))

    sc = ScrollableFrame(dashboard, bg=COLOR_BG)
    sc.pack(fill="both", expand=True)

    cards = sc.inner
    cards.configure(bg=COLOR_BG)
    cards.grid_columnconfigure(0, weight=1)
    cards.grid_columnconfigure(1, weight=1)

    for i, (btn_text, func) in enumerate(buttons_config):
        r, c2 = divmod(i, 2)
        card = Card(cards, bg=COLOR_CARD, border=COLOR_BORDER, hover=COLOR_LIGHT)
        card.grid(row=r, column=c2, sticky="nsew", padx=10, pady=10)

        inner = tk.Frame(card, bg=COLOR_CARD)
        inner.pack(fill="both", expand=True, padx=18, pady=18)

        tk.Label(inner, text=btn_text, font=FONT_CARD_TITLE, bg=COLOR_CARD, fg=COLOR_TEXT,
                 anchor="w", justify="left", wraplength=420).pack(fill="x", pady=(0, 12))

        make_button(inner, "Afficher →", lambda t=btn_text, f=func: run_action(t, f), bg=accent_color)\
            .pack(side="right")

    dashboard.tkraise()
    return page

# ============================================================
#  BUTTON CONFIGS
# ============================================================

ACP_BUTTONS = [
    ("📊 Statistiques descriptives", acp_stats),
    ("🔢 Matrice centrée-réduite", acp_cr),
    ("🎨 Heatmap corrélation (valeurs)", acp_corr),
    ("📈 Inerties & variance", acp_inerties),
    ("📍 Plan factoriel individus", acp_plan),
    ("⚫ Cercle de corrélation", acp_cercle),
    ("⭐ Qualité représentation (COS²)", acp_quality),
    ("🏆 Contributions", acp_contrib),
]

IA_BUTTONS = [
    ("👥 Visualiser clusters (contours)", ia_clusters),
    ("📊 Pourcentage par cluster", ia_pourcentage),
    ("🤖 Métriques Random Forest", ia_rf_metrics),
    ("🔮 Prédire nouvel individu", ia_predict_new),
]

AFC_BUTTONS = [
    ("📋 Tableau de contingence", afc_cont),
    ("📈 Matrice des fréquences", afc_freq),
    ("📉 Dépendance χ² & inertie", afc_dependance),
    ("📏 Distance χ²", afc_distance),
    ("📍 Plan factoriel AFC", afc_plan),
    ("✅ Test χ² & interprétation", afc_test),
]

CYBER_BUTTONS = [
    ("📁 Importer données (CYBER)", cyber_import),
    ("🌲 Isolation Forest", cyber_if),
    ("🎯 Local Outlier Factor (LOF)", cyber_lof),
    ("⚠️ Interprétation risques", cyber_interp),
]

# ============================================================
#  INIT APP
# ============================================================

def init_application():
    global root, container, k_var

    root = tk.Tk()
    root.title("Data Analyst Suite — Projet Final (Design Pro)")
    apply_global_style(root)
    root.geometry("1250x850")
    root.configure(bg=COLOR_BG)
    root.minsize(1050, 720)

    # ✅ k_var unique pour toutes les pages (fix)
    k_var = tk.StringVar(value=str(CURRENT_K))

    container = tk.Frame(root, bg=COLOR_BG)
    container.pack(fill="both", expand=True)
    container.grid_rowconfigure(0, weight=1)
    container.grid_columnconfigure(0, weight=1)

    create_start_page()
    create_menu_page()

    create_module_page("ACP", "ANALYSE EN COMPOSANTES PRINCIPALES (ACP)", ACP_BUTTONS, COLOR_BLUE)
    create_module_page("IA", "INTELLIGENCE ARTIFICIELLE — CLUSTERING & PRÉDICTION", IA_BUTTONS, COLOR_VIOLET)
    create_module_page("AFC", "ANALYSE FACTORIELLE DES CORRESPONDANCES (AFC)", AFC_BUTTONS, COLOR_GREEN)
    create_module_page("CYBER", "CYBERSÉCURITÉ — DÉTECTION D'ANOMALIES", CYBER_BUTTONS, COLOR_ORANGE)

    import_main_default(silent=True)

# ============================================================
#  MAIN
# ============================================================

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    init_application()
    show_frame("START")
    root.mainloop()
