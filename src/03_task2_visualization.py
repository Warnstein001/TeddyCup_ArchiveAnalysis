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