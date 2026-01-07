import os
import math
import zipfile
import pickle
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ------------
# unzip file
# ------------
date = 20260107
file_name = r"D:\研究\weekly"

zipfile_path = os.path.join(os.path.expanduser("~"), "Downloads", f"{date}_cs_newexp.zip")
with zipfile.ZipFile(zipfile_path, "r") as zf:
    zf.extractall(file_name)

data_dir = os.path.join(file_name, f"{date}_cs_newexp")
if not os.path.isdir(data_dir):
    cand = [
        os.path.join(file_name, d)
        for d in os.listdir(file_name)
        if os.path.isdir(os.path.join(file_name, d))
        and str(date) in d
        and "cs_newexp" in d
    ]
    if cand:
        data_dir = cand[0]
    else:
        raise FileNotFoundError(f"Cannot find extracted folder like {date}_cs_newexp under {file_name}")

# ------------
# load data
# ------------
batt_cycle_df = {}  
statis_df     = {}   
capacity_map  = {}   
cycle_map     = {}   

def loadfunction(i: int, j: int):
    cands = range(59, 64)  # 59:63
    fns = [os.path.join(data_dir, f"240076-{i}-{j}-28185739{c}.xlsx") for c in cands]
    existing = [fn for fn in fns if os.path.exists(fn)]
    if not existing:
        raise FileNotFoundError("Cannot find any file among: " + ", ".join(fns))
    N = existing[0]

    df_step  = pd.read_excel(N, sheet_name="step", engine="openpyxl")
    df_cycle = pd.read_excel(N, sheet_name="cycle", engine="openpyxl")

    cp_idx = df_step["工步類型"].astype(str).eq("恆流放電")
    cap = df_step.loc[cp_idx, "容量(Ah)"].dropna().to_numpy()

    k = 10 * i + j
    batt_cycle_df[k] = df_cycle
    statis_df[k]     = df_step
    capacity_map[k]  = cap
    cycle_map[k]     = int(len(cap))

for i in range(1, 5):
    for j in range(1, 9):
        loadfunction(i, j)
        print(f"Battery {i}-{j} Done")

# -------------------
# generate report
# -------------------
report_dir = os.path.join(data_dir, "report")
os.makedirs(report_dir, exist_ok=True)
os.chdir(report_dir)

i_list = list(range(1, 5))
ja = list(range(1, 9, 2))  # 1,3,5,7
jb = list(range(2, 9, 2))  # 2,4,6,8

def build_plan_indices(i_list, j_list):
    idx_mat = np.array([[10 * ii + jj for ii in i_list] for jj in j_list])
    idx_vec = idx_mat.flatten(order="F")
    return idx_vec

idxA = build_plan_indices(i_list, ja)
idxB = build_plan_indices(i_list, jb)

PlanA_cycle = [cycle_map[k] for k in idxA]
PlanB_cycle = [cycle_map[k] for k in idxB]

PlanA_capacity = [capacity_map[k] for k in idxA]
PlanB_capacity = [capacity_map[k] for k in idxB]

PlanA_standcapacity = [cap / cap[0] for cap in PlanA_capacity]
PlanB_standcapacity = [cap / cap[0] for cap in PlanB_capacity]

battery_name_planA = [f"Battery {k//10}-{k%10}" for k in idxA]
battery_name_planB = [f"Battery {k//10}-{k%10}" for k in idxB]

PlanA_dat_plot = dict(standcapacity=PlanA_standcapacity, cycle=PlanA_cycle, Battery_name=battery_name_planA)
PlanB_dat_plot = dict(standcapacity=PlanB_standcapacity, cycle=PlanB_cycle, Battery_name=battery_name_planB)

xlim_num = [min(PlanA_cycle), min(PlanB_cycle)]

def overall_ylim(standcaps, xlim):
    mins, maxs = [], []
    for arr in standcaps:
        seg = arr[:xlim]
        if len(seg) == 0:
            continue
        mins.append(float(np.min(seg)))
        maxs.append(float(np.max(seg)))
    return (min(mins) - 0.01, max(maxs) + 0.01)

cap_range = [overall_ylim(PlanA_standcapacity, xlim_num[0]),
             overall_ylim(PlanB_standcapacity, xlim_num[1])]

state = {
    "date": date,
    "data_dir": data_dir,
    "report_dir": report_dir,
    "batt_cycle_df": batt_cycle_df,
    "statis_df": statis_df,
    "capacity_map": capacity_map,
    "cycle_map": cycle_map,
    "PlanA_dat_plot": PlanA_dat_plot,
    "PlanB_dat_plot": PlanB_dat_plot,
    "xlim_num": xlim_num,
    "cap_range": cap_range,
    "idxA": idxA,
    "idxB": idxB,
}
with open(f"{date}_cs_newexp.pkl", "wb") as f:
    pickle.dump(state, f)

def make_col_pattern(xlim, mode):
    if mode == "each":
        base = np.repeat([0, 1, 2], 20)
    elif mode == "times":
        base = np.tile([0, 1, 2], 20)
    else:
        raise ValueError("mode must be 'each' or 'times'")
    rep = int(math.ceil(xlim / len(base)))
    return np.tile(base, rep)[:xlim]

def plot_plan(out_png, plan, xlim, ylim, vline, col_mode):
    fig, axes = plt.subplots(4, 4, figsize=(12, 12), dpi=600)
    axes = axes.flatten()

    col_vals = make_col_pattern(xlim, col_mode)

    for idx in range(16):
        ax = axes[idx]
        y = plan["standcapacity"][idx][:xlim]
        x = np.arange(1, len(y) + 1)

        ax.scatter(x, y, c=col_vals[:len(y)], cmap="tab10", s=14)
        ax.set_xlim(0, xlim)
        ax.set_ylim(ylim[0], ylim[1])
        ax.set_xlabel("cycle", fontsize=14)
        ax.set_ylabel("Standard Capacity", fontsize=14)
        ax.set_title(plan["Battery_name"][idx], fontsize=14)
        ax.axvline(vline, color="gray", linewidth=1)
        ax.tick_params(axis="both", labelsize=12)

    plt.tight_layout()
    fig.savefig(out_png)
    plt.close(fig)

plot_plan("PlanA_capacity.png", PlanA_dat_plot, xlim_num[0], cap_range[0], vline=28, col_mode="each")
plot_plan("PlanB_capacity.png", PlanB_dat_plot, xlim_num[1], cap_range[1], vline=31, col_mode="times")
