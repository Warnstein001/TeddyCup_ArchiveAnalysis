# src/03_task2_visualization.py

import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
import importlib.util
import os

from utils import calc_work_hours   # 虽然 2.1 用不到，但后面 2.2 会用


# ======================
# 动态导入 01_preprocess.py
# ======================
spec = importlib.util.spec_from_file_location(
    "preprocess",
    os.path.join(os.path.dirname(__file__), "01_preprocess.py")
)
preprocess_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(preprocess_module)
preprocess_data = preprocess_module.preprocess_data


# ======================
# Task 2.1
# ======================
def plot_task2_1_daily_finished_count(df):
    """
    每天 × 工序 完成案卷数量（簇状柱状图）
    输出：result/figures/task2_1.png
    """

    # ---------- Step 1：只保留完成记录 ----------
    df_finished = df[df["is_finished"]].copy()

    # ---------- Step 2：提取完成日期 ----------
    # 使用完成节点时间 dNODE_TIME
    df_finished["date"] = df_finished["dNODE_TIME"].dt.date

    # ---------- Step 3：按 日期 × 工序 统计完成案卷数 ----------
    daily_count = (
        df_finished
        .groupby(["date", "工序"])["sARCH_ID"]
        .nunique()
        .reset_index(name="completed_cases")
    )

    # ---------- Step 4：工序名称映射为英文（防乱码） ----------
    process_map = {
        "扫描": "Scanning",
        "图像处理": "Image Processing",
        "自检全检": "Inspection",
        "PDF处理": "PDF Generation"
    }

    daily_count["Process"] = daily_count["工序"].map(process_map)

    # ---------- Step 5：转换为透视表 ----------
    pivot_data = (
        daily_count
        .pivot(index="date", columns="Process", values="completed_cases")
        .fillna(0)
    )

    # ---------- Step 6：作图 ----------
    plt.figure(figsize=(12, 6))
    pivot_data.plot(kind="bar", width=0.8)

    plt.xlabel("Date")
    plt.ylabel("Completed Cases")
    plt.title("Daily Completed Cases by Process")
    plt.legend(title="Process")
    plt.xticks(rotation=45)
    plt.tight_layout()

    # ---------- Step 7：保存图片 ----------
    output_dir = Path("result/figures")
    output_dir.mkdir(parents=True, exist_ok=True)

    plt.savefig(
        output_dir / "task2_1.png",
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()


# ======================
# Task 2.2
# ======================
def plot_task2_2_daily_workload(df):
    """
    每天 × 工序 投入工作量（人·小时）
    输出：result/figures/task2_2.png
    """

    # ---------- Step 1：只保留完成记录 ----------
    df_finished = df[df["is_finished"]].copy()

    # ---------- Step 2：按 工序 × 批次 聚合时间区间 ----------
    batch_time = (
        df_finished
        .groupby(["工序", "sBatch_number"])
        .agg(
            batch_start=("dUPDATE_TIME", "min"),
            batch_end=("dNODE_TIME", "max")
        )
        .reset_index()
    )

    # ---------- Step 3：计算每个批次的有效工作时长 ----------
    batch_time["work_hours"] = batch_time.apply(
        lambda row: calc_work_hours(
            row["batch_start"],
            row["batch_end"]
        ),
        axis=1
    )

    # ---------- Step 4：确定批次所属日期（用开始日期） ----------
    batch_time["date"] = batch_time["batch_start"].dt.date

    # ---------- Step 5：工序名称映射为英文 ----------
    process_map = {
        "扫描": "Scanning",
        "图像处理": "Image Processing",
        "自检全检": "Inspection",
        "PDF处理": "PDF Generation"
    }
    batch_time["Process"] = batch_time["工序"].map(process_map)

    # ---------- Step 6：按 日期 × 工序 汇总工作量 ----------
    daily_workload = (
        batch_time
        .groupby(["date", "Process"])["work_hours"]
        .sum()
        .reset_index(name="Workload (Person-Hours)")
    )

    # ---------- Step 7：转换为透视表 ----------
    pivot_data = (
        daily_workload
        .pivot(index="date", columns="Process", values="Workload (Person-Hours)")
        .fillna(0)
    )

    # ---------- Step 8：作图 ----------
    plt.figure(figsize=(12, 6))
    pivot_data.plot(kind="bar", width=0.8)

    plt.xlabel("Date")
    plt.ylabel("Workload (Person-Hours)")
    plt.title("Daily Workload by Process")
    plt.legend(title="Process")
    plt.xticks(rotation=45)
    plt.tight_layout()

    # ---------- Step 9：保存图片 ----------
    output_dir = Path("result/figures")
    output_dir.mkdir(parents=True, exist_ok=True)

    plt.savefig(
        output_dir / "task2_2.png",
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

# ======================
# Task 2.3
# ======================
def plot_task2_3_daily_rework_ratio(df):
    """
    每天 × 工序 返工占比（堆积面积图）
    输出：result/figures/task2_3.png
    """

    # ---------- Step 1：只保留完成记录 ----------
    df_finished = df[df["is_finished"]].copy()

    # ---------- Step 2：提取完成日期 ----------
    df_finished["date"] = df_finished["dNODE_TIME"].dt.date

    # ---------- Step 3：统计每天 × 工序 完成案卷数（分母） ----------
    total_daily = (
        df_finished
        .groupby(["date", "工序"])["sARCH_ID"]
        .nunique()
        .reset_index(name="total_cases")
    )

    # ---------- Step 4：统计每天 × 工序 返工案卷数（分子） ----------
    rework_daily = (
        df_finished[df_finished["is_rework"]]
        .groupby(["date", "工序"])["sARCH_ID"]
        .nunique()
        .reset_index(name="rework_cases")
    )

    # ---------- Step 5：合并并计算返工占比 ----------
    daily_ratio = pd.merge(
        total_daily,
        rework_daily,
        on=["date", "工序"],
        how="left"
    )

    daily_ratio["rework_cases"] = daily_ratio["rework_cases"].fillna(0)
    daily_ratio["rework_ratio"] = (
        daily_ratio["rework_cases"] / daily_ratio["total_cases"]
    )

    # ---------- Step 6：工序名称映射为英文 ----------
    process_map = {
        "扫描": "Scanning",
        "图像处理": "Image Processing",
        "自检全检": "Inspection",
        "PDF处理": "PDF Generation"
    }
    daily_ratio["Process"] = daily_ratio["工序"].map(process_map)

    # ---------- Step 7：转换为透视表 ----------
    pivot_data = (
        daily_ratio
        .pivot(index="date", columns="Process", values="rework_ratio")
        .fillna(0)
        .sort_index()
    )

    # ---------- Step 8：作堆积面积图 ----------
    plt.figure(figsize=(12, 6))
    pivot_data.plot.area(alpha=0.8)

    plt.xlabel("Date")
    plt.ylabel("Rework Ratio")
    plt.title("Daily Rework Ratio by Process")
    plt.legend(title="Process", loc="upper left")
    plt.tight_layout()

    # ---------- Step 9：保存图片 ----------
    output_dir = Path("result/figures")
    output_dir.mkdir(parents=True, exist_ok=True)

    plt.savefig(
        output_dir / "task2_3.png",
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

# ======================
# Task 2.4
# ======================
def plot_task2_4_image_user_rework_pie(df, top_n=8):
    """
    图像处理工序 —— 操作人员返工占比（饼图）
    输出：result/figures/task2_4.png
    """

    # ---------- Step 1：只保留完成记录 ----------
    df_finished = df[df["is_finished"]].copy()

    # ---------- Step 2：限定工序为“图像处理” ----------
    df_img = df_finished[df_finished["工序"] == "图像处理"].copy()

    # ---------- Step 3：统计每个操作人员的完成案卷数（分母） ----------
    total_cases = (
        df_img
        .groupby("iUSER_ID")["sARCH_ID"]
        .nunique()
        .reset_index(name="total_cases")
    )

    # ---------- Step 4：统计每个操作人员的返工案卷数（分子） ----------
    rework_cases = (
        df_img[df_img["is_rework"]]
        .groupby("iUSER_ID")["sARCH_ID"]
        .nunique()
        .reset_index(name="rework_cases")
    )

    # ---------- Step 5：合并并计算返工占比 ----------
    user_ratio = pd.merge(
        total_cases,
        rework_cases,
        on="iUSER_ID",
        how="left"
    )

    user_ratio["rework_cases"] = user_ratio["rework_cases"].fillna(0)
    user_ratio["rework_ratio"] = user_ratio["rework_cases"] / user_ratio["total_cases"]

    # ---------- Step 6：按返工案卷数排序，取 Top N ----------
    user_ratio_sorted = user_ratio.sort_values(
        by="rework_cases",
        ascending=False
    )

    top_users = user_ratio_sorted.head(top_n)
    others = user_ratio_sorted.iloc[top_n:]

    # 合并 Others
    if not others.empty:
        others_row = pd.DataFrame({
            "iUSER_ID": ["Others"],
            "total_cases": [others["total_cases"].sum()],
            "rework_cases": [others["rework_cases"].sum()],
            "rework_ratio": [others["rework_cases"].sum() / max(others["total_cases"].sum(), 1)]
        })
        pie_data = pd.concat([top_users, others_row], ignore_index=True)
    else:
        pie_data = top_users.copy()

    # ---------- Step 7：准备饼图数据 ----------
    labels = pie_data["iUSER_ID"].astype(str)
    sizes = pie_data["rework_cases"]

    # 如果全为 0，避免报错
    if sizes.sum() == 0:
        print("No rework cases found in Image Processing. Pie chart not generated.")
        return

    # ---------- Step 8：作图 ----------
    plt.figure(figsize=(8, 8))
    plt.pie(
        sizes,
        labels=labels,
        autopct="%1.1f%%",
        startangle=140
    )
    plt.title("Rework Distribution by Operator (Image Processing)")
    plt.axis("equal")  # 保证为正圆

    # ---------- Step 9：保存图片 ----------
    output_dir = Path("result/figures")
    output_dir.mkdir(parents=True, exist_ok=True)

    plt.savefig(
        output_dir / "task2_4.png",
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()



# ======================
# main
# ======================
if __name__ == "__main__":
    print("Loading and preprocessing data...")
    df = preprocess_data("data/data.xlsx")

    print("Generating Task 2.1 figure...")
    plot_task2_1_daily_finished_count(df)

    print("Generating Task 2.2 figure...")
    plot_task2_2_daily_workload(df)

    print("Task 2 figures saved to result/figures/")

    print("Generating Task 2.3 figure...")
    plot_task2_3_daily_rework_ratio(df)
    
    print("Generating Task 2.4 figure...")
    plot_task2_4_image_user_rework_pie(df, top_n=8)
