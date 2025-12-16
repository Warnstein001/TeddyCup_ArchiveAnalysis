# src/04_task3_pattern.py

import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from utils import calc_work_hours
import importlib.util
import os

# 动态导入 preprocess
spec = importlib.util.spec_from_file_location(
    "preprocess",
    os.path.join(os.path.dirname(__file__), "01_preprocess.py")
)
preprocess_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(preprocess_module)
preprocess_data = preprocess_module.preprocess_data


# ======================
# Task 3.1
# ======================
def analyze_processing_time_distribution(df):
    """
    Task 3.1: Distribution of processing time
    """

    # ---------- Step 1：只保留有效完成记录 ----------
    df_valid = df[df["is_finished"]].copy()

    # ---------- Step 2：计算领取-提交工作时长 ----------
    df_valid["processing_hours"] = df_valid.apply(
        lambda row: calc_work_hours(
            row["dUPDATE_TIME"],
            row["dNODE_TIME"]
        ),
        axis=1
    )

    # 去除异常
    df_valid = df_valid[df_valid["processing_hours"] > 0]

    # ---------- Step 3：绘制分布图 ----------
    plt.figure(figsize=(8, 5))
    plt.hist(
        df_valid["processing_hours"],
        bins=50,
        density=True,
        alpha=0.75
    )

    plt.xlabel("Processing Time (hours)")
    plt.ylabel("Density")
    plt.title("Distribution of Processing Time (Receive → Submit)")
    plt.tight_layout()

    # ---------- Step 4：保存 ----------
    output_dir = Path("result/figures")
    output_dir.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_dir / "task3_1_processing_time_dist.png", dpi=300)
    plt.close()



if __name__ == "__main__":
    df = preprocess_data("data/data.xlsx")

    print("Running Task 3.1...")
    analyze_processing_time_distribution(df)
