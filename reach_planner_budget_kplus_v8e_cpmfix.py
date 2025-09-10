
import os
from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# =======================================
# i.com Reach Planner — v8e (CPM fix)
#  - Методология = v8d (базовая), НО CPM из сайдбара реально участвует в расчётах
#  - Defaults: ExtractJuly24-Jul25_fin.xlsx / Sheet2
#  - Caps масштабируются по аудитории
#  - Alpha: взвешенная медиана по строкам со шринкажем к глобальному (m0=20)
#  - TOTAL: powered-product среднее с нижней отсечкой (>= max per-site)
#  - Cache-busting по mtime Excel (подхватывает обновления файла с тем же именем)
# =======================================

st.set_page_config(page_title="i.com Reach Planner", layout="wide")
st.title("i.com Reach Planner")

# -------- Вселенные (people) --------
FIXED_UNIVERSE = {
    "All18+": 84400000,
    "All25-45": 45930000,
    "All35-55": 44950000,
    "M25-45": 22628000,
    "M35-55": 21611000,
    "W18-35": 16210000,
    "W25-45": 23302000,
    "W35-55": 23338000,
    "W25-65": 44699000,
}

# -------- Каппинги (All18+; PEOPLE) --------
CAP_OVERRIDES = {
    "astralab": 19000000, "auto": 30000000, "autonwsru": 16000000, "avito": 72000000,
    "betweenx": 27000000, "buzzoola": 20000000, "byyd": 20000000, "dgtlall": 65000000,
    "everest": 5000000, "getintent": 88000000, "gismeteo": 21000000, "gpm": 81000000,
    "hybrid": 66000000, "interpool": 10000000, "ivi": 13100000, "kinopoisk": 15000000,
    "kommsnt": 5000000, "lenta": 13000000, "link": 20000000, "mediasniper": 51000000,
    "mobidriven": 18000000, "mobx": 20000000, "mts": 61000000, "mytarget": 80000000,
    "nativerent": 10000000, "otm": 53000000, "ozon": 81000000, "pkauto": 25000000,
    "pknews": 25000000, "pksmtv": 10000000, "rambler": 33000000, "rbc": 16000000,
    "reddigital": 15000000, "redlama": 25000000, "sjp": 30000000, "soloway": 46000000,
    "solta": 27000000, "targetnative": 10000000, "videoint": 15000000, "vk": 102000000,
    "yandex": 102000000,
}

# -------- Загрузка данных --------
DEFAULT_NAME = "ExtractJuly24-Jul25_fin.xlsx"
DEFAULT_SHEET = "Sheet2"
DATA_PATH = Path(os.environ.get("REACH_DATA_PATH", DEFAULT_NAME))
SHEET_NAME = os.environ.get("REACH_XLS_SHEET", DEFAULT_SHEET)

def _normalize_audience(a: str) -> str:
    s = str(a).strip().replace("—","-").replace("–","-").replace(" ","")
    if s.startswith("М"): s = "M" + s[1:]
    if s.startswith("Ж"): s = "W" + s[1:]
    if s.lower().startswith("all"): s = "All" + s[3:]
    return s

def _universe(aud_label: str) -> float:
    return float(FIXED_UNIVERSE.get(_normalize_audience(aud_label), 0.0))

@st.cache_data(show_spinner=True)
def load_df_cached(path_str: str, sheet_name: str, cache_bust: float) -> pd.DataFrame:
    """Кэш включает mtime файла -> обновления XLSX подхватываются автоматически."""
    path = Path(path_str)
    if not path.is_absolute():
        path = Path.cwd() / path
    xl = pd.ExcelFile(path)
    if sheet_name not in xl.sheet_names:
        sheet_name = xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    if "Аудитория" in df.columns:
        df["Аудитория"] = df["Аудитория"].map(_normalize_audience)
    return df

# cache-bust token
try:
    mtime = (Path.cwd() / DATA_PATH).stat().st_mtime if not Path(DATA_PATH).is_absolute() else Path(DATA_PATH).stat().st_mtime
except Exception:
    mtime = 0.0

try:
    df = load_df_cached(str(DATA_PATH), str(SHEET_NAME), mtime)
except Exception as e:
    st.error(f"Ошибка чтения файла данных: {e}")
    st.stop()

# Найти столбцы N..W (10 ступеней частоты)
def excel_col_letter(idx0:int) -> str:
    letters = ""
    idx = idx0 + 1
    while idx:
        idx, rem = divmod(idx-1, 26)
        letters = chr(65+rem) + letters
    return letters

letters_map = {excel_col_letter(i): name for i, name in enumerate(df.columns)}
STEP_LETTERS = ["N","O","P","Q","R","S","T","U","V","W"]
STEP_COLS = [letters_map[x] for x in STEP_LETTERS if x in letters_map]
if len(STEP_COLS) != 10:
    st.error("Не найдены столбцы распределения показов N–W (10 штук).")
    st.stop()

# Текстовые поля
for c in ["Сайт", "Формат", "Аудитория"]:
    if c in df.columns:
        df[c] = df[c].astype(str).str.strip()

# ===== Сайдбар =====
st.sidebar.header("Фильтры")
sites_all = sorted(df["Сайт"].dropna().unique().tolist(), key=lambda x: str(x).lower())
default_sites = sites_all[:1]  # одна площадка по умолчанию
sites_sel = st.sidebar.multiselect("Площадки", sites_all, default=default_sites)

df_sites = df[df["Сайт"].isin(sites_sel)] if sites_sel else df.copy()
formats_opts = sorted(df_sites["Формат"].dropna().unique().tolist(), key=lambda x: str(x).lower())
fmt_sel = st.sidebar.selectbox("Формат", formats_opts)

# Аудитории = пересечение по выбранным площадкам в выбранном формате
if fmt_sel:
    aud_sets = []
    for s in (sites_sel if sites_sel else sites_all):
        seg = df[(df["Сайт"] == s) & (df["Формат"].str.lower() == fmt_sel.lower())]
        aud_s = set(seg["Аудитория"].dropna().unique().tolist())
        if aud_s:
            aud_sets.append(aud_s)
    if aud_sets:
        aud_common = set.intersection(*aud_sets) if len(aud_sets) > 1 else aud_sets[0]
    else:
        aud_common = set()
else:
    aud_common = set()

aud_opts = sorted(list(aud_common))
aud_sel = st.sidebar.selectbox("Аудитория", aud_opts)

def _U(): return _universe(aud_sel) if aud_sel else 0.0
U = _U()

# Экономика и цели
st.sidebar.header("Экономика и цель")
k_all = list(range(1, 11))
k1 = st.sidebar.selectbox("Частота A (k+)", k_all, index=0)
k2 = st.sidebar.selectbox("Частота B (k+)", [k for k in k_all if k != k1], index=1 if 2 != k1 else 0)

mode = st.sidebar.radio("Режим бюджета", ["Общий бюджет кампании", "Бюджет по площадкам"], index=0)

site_params = {}
for s in sites_sel:
    st.sidebar.subheader(s)
    cpm = st.sidebar.number_input(f"CPM для {s}, ₽", min_value=1.0, value=200.0, step=1.0, key=f"cpm_{s}")
    site_params[s] = {"cpm": cpm}

total_budget = None
percs = {}
if mode == "Общий бюджет кампании":
    total_budget = st.sidebar.number_input("Общий бюджет, ₽", min_value=0.0, value=10_000_000.0, step=500_000.0)
    manual_split = st.sidebar.checkbox("Задать доли бюджета по площадкам, %", value=False)
    if manual_split:
        default = round(100.0/max(len(sites_sel),1), 2) if sites_sel else 0.0
        for s in sites_sel:
            percs[s] = st.sidebar.number_input(f"{s}, %", min_value=0.0, max_value=100.0, value=default, step=1.0, key=f"pct_{s}")
        if sum(percs.values()) <= 0 and len(sites_sel) > 0:
            for s in sites_sel: percs[s] = 100.0/len(sites_sel)
    else:
        if len(sites_sel) > 0:
            for s in sites_sel: percs[s] = 100.0/len(sites_sel)
else:
    for s in sites_sel:
        bud = st.sidebar.number_input(f"Бюджет {s}, ₽", min_value=0.0, value=0.0, step=100000.0, key=f"bud_{s}")
        site_params[s]["budget"] = bud

st.sidebar.divider()
y_units = st.sidebar.radio("Единицы по Y", ["%", "млн чел."], index=1 if U>0 else 0)
max_budget_for_plot = st.sidebar.number_input("Максимум бюджета на оси X, ₽", min_value=0.0, value=(total_budget or 20_000_000.0), step=1_000_000.0)
num_points = st.sidebar.slider("Число точек на оси бюджета", min_value=20, max_value=200, value=80)

# ===== Математика =====
def proportional_match_to_E(step_vals_row: np.ndarray, E: float) -> np.ndarray:
    s = float(step_vals_row.sum())
    if E <= 0:
        return np.zeros_like(step_vals_row)
    if s <= 0:
        return np.full_like(step_vals_row, E/len(step_vals_row))
    return step_vals_row * (E / s)

def people_shape_w_from_steps(adj_steps_row: np.ndarray) -> np.ndarray:
    s_k = adj_steps_row / max(adj_steps_row.sum(), 1e-12)
    k = np.arange(1, len(s_k)+1, dtype=float)
    w = s_k / k
    wsum = w.sum()
    return w / wsum if wsum>0 else np.full_like(w, 1.0/len(w))

def weighted_median(values: np.ndarray, weights: np.ndarray) -> float:
    order = np.argsort(values)
    v = values[order]; w = weights[order]
    cw = (w.cumsum()/w.sum()) if w.sum()>0 else np.zeros_like(w)
    idx = np.searchsorted(cw, 0.5)
    return float(v[min(idx, len(v)-1)])

def get_scaled_cap_people(site_name: str, aud_label: str) -> float | None:
    key = str(site_name).strip().lower()
    cap_all = CAP_OVERRIDES.get(key)
    if cap_all is None:
        return None
    U_all = FIXED_UNIVERSE.get("All18+", 0.0)
    U_a = _universe(aud_label)
    if U_all and U_a:
        return float(cap_all) * (U_a / U_all)
    return float(cap_all)

@st.cache_data(show_spinner=False)
def build_site_model(site: str, fmt: str, aud: str, df: pd.DataFrame, step_cols: list[str], m0:int=20) -> dict | None:
    seg = df[(df["Сайт"]==site) & (df["Формат"].str.lower()==fmt.lower()) & (df["Аудитория"]==aud)].copy()
    if seg.empty: return None
    seg[step_cols] = seg[step_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    seg["Показы"] = pd.to_numeric(seg["Показы"], errors="coerce").fillna(0.0)
    seg["Охват"] = pd.to_numeric(seg["Охват"], errors="coerce").fillna(0.0)
    U_loc = _universe(aud)
    if U_loc <= 0: return None

    ws, alphas, weights = [], [], []
    for _, r in seg.iterrows():
        steps = r[step_cols].to_numpy(dtype=float)
        E = float(r["Показы"]); R1 = float(r["Охват"])
        adj = proportional_match_to_E(steps, E)
        w = people_shape_w_from_steps(adj); ws.append(w); weights.append(E)
        mu = E / U_loc if U_loc>0 else 0.0
        r1 = np.clip(R1 / U_loc, 1e-9, 0.999999) if U_loc>0 else 0.0
        alpha = (-np.log(1.0 - r1) / mu) if mu > 0 else 0.0
        alphas.append(alpha)

    alphas = np.array(alphas); weights = np.array(weights); ws = np.array(ws)
    # site alpha = E-weighted median
    order = np.argsort(alphas); v = alphas[order]; w = weights[order]
    cw = (w.cumsum()/w.sum()) if w.sum()>0 else np.zeros_like(w)
    idx = np.searchsorted(cw, 0.5)
    alpha_site = float(v[min(idx, len(v)-1)]) if len(v)>0 else 0.0
    n = int(len(seg))

    # global alpha for same (fmt,aud) across all sites
    seg_all = df[(df["Формат"].str.lower()==fmt.lower()) & (df["Аудитория"]==aud)].copy()
    seg_all["Показы"] = pd.to_numeric(seg_all["Показы"], errors="coerce").fillna(0.0)
    seg_all["Охват"] = pd.to_numeric(seg_all["Охват"], errors="coerce").fillna(0.0)
    alphas_all, weights_all = [], []
    for _, r in seg_all.iterrows():
        E=float(r["Показы"]); R1=float(r["Охват"])
        mu = E / U_loc if U_loc>0 else 0.0
        r1 = np.clip(R1 / U_loc, 1e-9, 0.999999) if U_loc>0 else 0.0
        a = (-np.log(1.0 - r1) / mu) if mu>0 else 0.0
        alphas_all.append(a); weights_all.append(E)
    if len(alphas_all) > 0:
        aa = np.array(alphas_all); ww = np.array(weights_all)
        order = np.argsort(aa); aa = aa[order]; ww = ww[order]
        cw = (ww.cumsum()/ww.sum()) if ww.sum()>0 else np.zeros_like(ww)
        idx = np.searchsorted(cw, 0.5)
        alpha_global = float(aa[min(idx, len(aa)-1)])
    else:
        alpha_global = 0.25

    # shrinkage
    wshr = n / (n + m0)
    alpha_blend = wshr * alpha_site + (1 - wshr) * alpha_global

    w_site = np.average(ws, axis=0, weights=weights if weights.sum()>0 else None)
    return {"site": site, "fmt": fmt, "aud": aud, "U": U_loc, "w": w_site, "alpha": float(alpha_blend)}

# Распределение бюджета
def budgets_for_total(x_total: float, sites: list[str]) -> dict[str,float]:
    if not sites: return {}
    if mode == "Общий бюджет кампании":
        total_pct = sum(percs.values()) if percs else 0.0
        if total_pct <= 0:
            return {s: x_total/len(sites) for s in sites}
        return {s: x_total * (percs[s]/total_pct) for s in sites}
    else:
        base = sum(site_params[s].get("budget",0.0) for s in sites)
        if base <= 0:
            return {s: x_total/len(sites) for s in sites}
        scale = x_total / base
        return {s: site_params[s]["budget"] * scale for s in sites}

# Reach engine
def reach_kplus(model: dict, cpm_rub: float, budget_rub: float, k_plus: int, aud: str) -> float:
    U_loc = model["U"]; alpha = model["alpha"]; w = model["w"]
    if U_loc <= 0 or alpha <= 0: return 0.0
    mu = (1000.0 * (budget_rub / cpm_rub)) / U_loc
    cap_people = get_scaled_cap_people(model["site"], aud)
    if cap_people and cap_people > 0:
        phi = min(cap_people / U_loc, 0.999999)
        r1 = phi * (1.0 - np.exp(-alpha * mu / phi))
    else:
        r1 = 1.0 - np.exp(-alpha * mu)
    if k_plus <= 1:
        return U_loc * r1
    p_exact = r1 * w
    idx = max(1, min(int(k_plus), len(w)))
    return U_loc * p_exact[idx-1:].sum()

def total_avg_powered_lower_bounded(models: dict[str,dict], cpms: dict[str,float], budgets: dict[str,float], k_plus: int, aud: str) -> float:
    if not models: return 0.0
    U_loc = list(models.values())[0]["U"]
    ps = []
    for s,m in models.items():
        cpm = cpms.get(s, 200.0)
        p = reach_kplus(m, cpm, budgets.get(s,0.0), k_plus, aud) / U_loc
        ps.append(p)
    ps = np.clip(np.array(ps), 0.0, 0.999999)
    max_p = float(ps.max()) if len(ps)>0 else 0.0
    q_prod = float(np.prod(1.0 - ps)) if len(ps)>0 else 1.0
    sum_p = float(ps.sum()); n=len(ps)
    def gamma(beta,eta): return 1.0/(1.0 + beta*max(n-1,0)*(sum_p**eta))
    p1 = 1.0 - (q_prod ** gamma(1.0, 1.5))
    p2 = 1.0 - (q_prod ** gamma(1.5, 1.0))
    pav = max(0.5*(p1+p2), max_p)
    return pav * U_loc

# Построение моделей по выбранным площадкам
models = {}
for s in sites_sel:
    m = build_site_model(s, fmt_sel, aud_sel, df, STEP_COLS, m0=20)
    if m is not None: models[s] = m

# CPM из сайдбара
cpms = {s: site_params.get(s, {}).get("cpm", 200.0) for s in models.keys()}

# ===== График =====
fig, ax = plt.subplots(figsize=(10,6))
xs = np.linspace(0.0, max_budget_for_plot, int(num_points))

def to_unit(people: float) -> float:
    return (people/1e6) if y_units=="млн чел." else (people/(U if U>0 else 1.0)*100.0)

for s,m in models.items():
    ys1, ys2 = [], []
    for x in xs:
        b = budgets_for_total(x, list(models.keys()))
        y1 = reach_kplus(m, cpms.get(s,200.0), b.get(s,0.0), k1, aud_sel); ys1.append(to_unit(y1))
        y2 = reach_kplus(m, cpms.get(s,200.0), b.get(s,0.0), k2, aud_sel); ys2.append(to_unit(y2))
    ax.plot(xs/1e6, ys1, marker="o", label=f"{s} — {k1}+")
    ax.plot(xs/1e6, ys2, linestyle="--", label=f"{s} — {k2}+")

# TOTAL
ys_t1, ys_t2 = [], []
for x in xs:
    b = budgets_for_total(x, list(models.keys()))
    y1 = total_avg_powered_lower_bounded(models, cpms, b, k1, aud_sel)
    y2 = total_avg_powered_lower_bounded(models, cpms, b, k2, aud_sel)
    ys_t1.append(to_unit(y1)); ys_t2.append(to_unit(y2))
ax.plot(xs/1e6, ys_t1, linewidth=3, label=f"TOTAL — {k1}+")
ax.plot(xs/1e6, ys_t2, linewidth=3, linestyle="--", label=f"TOTAL — {k2}+")

ax.set_xlabel("Бюджет, млн ₽"); ax.set_ylabel(f"Reach ({y_units})")
if aud_sel: ax.set_title(f"{aud_sel} • {k1}+ / {k2}+ • {fmt_sel}")
ax.grid(True); ax.legend()
st.pyplot(fig)

# ===== Таблица =====
st.subheader("Расчёт на выбранном бюджете")
chosen_total = (total_budget if mode=="Общий бюджет кампании" else sum(site_params[s].get("budget",0.0) for s in sites_sel))
b = budgets_for_total(chosen_total, list(models.keys()))
rows = []
for s,m in models.items():
    y1 = reach_kplus(m, cpms.get(s,200.0), b.get(s,0.0), k1, aud_sel)
    y2 = reach_kplus(m, cpms.get(s,200.0), b.get(s,0.0), k2, aud_sel)
    rows.append({"Площадка": s, "CPM, ₽": cpms.get(s,200.0), "Бюджет, ₽": b.get(s,0.0),
                 f"Reach {k1}+ ({y_units})": round(to_unit(y1),2),
                 f"Reach {k2}+ ({y_units})": round(to_unit(y2),2)})
# TOTAL row
y1_tot = total_avg_powered_lower_bounded(models, cpms, b, k1, aud_sel)
y2_tot = total_avg_powered_lower_bounded(models, cpms, b, k2, aud_sel)
rows.append({"Площадка":"TOTAL СРЕДНЕЕ (рекомендуемое)", "CPM, ₽":"", "Бюджет, ₽": sum(b.values()),
             f"Reach {k1}+ ({y_units})": round(to_unit(y1_tot),2),
             f"Reach {k2}+ ({y_units})": round(to_unit(y2_tot),2)})
st.dataframe(pd.DataFrame(rows))
