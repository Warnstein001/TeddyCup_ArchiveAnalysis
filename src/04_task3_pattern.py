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

# ======================
# Task 3.2
# ======================
def cluster_operator_behavior(df, k=3):
    """
    Task 3.2: Operator behavior clustering
    """

    from sklearn.cluster import KMeans
    from sklearn.preprocessing import StandardScaler

    # ---------- Step 1：只保留完成记录 ----------
    df_valid = df[df["is_finished"]].copy()

    # ---------- Step 2：计算领取-提交工作时长 ----------
    df_valid["processing_hours"] = df_valid.apply(
        lambda row: calc_work_hours(
            row["dUPDATE_TIME"],
            row["dNODE_TIME"]
        ),
        axis=1
    )
    df_valid = df_valid[df_valid["processing_hours"] > 0]

    # ---------- Step 3：构建人员级特征 ----------
    def long_task_ratio(x):
        return (x > x.quantile(0.9)).mean()

    features = (
        df_valid
        .groupby("iUSER_ID")["processing_hours"]
        .agg(
            avg_time="mean",
            median_time="median",
            p90_time=lambda x: x.quantile(0.9),
            long_ratio=long_task_ratio,
            case_count="count"
        )
        .reset_index()
    )

    # ---------- Step 4：标准化 ----------
    feature_cols = ["avg_time", "median_time", "p90_time", "long_ratio", "case_count"]
    scaler = StandardScaler()
    X = scaler.fit_transform(features[feature_cols])

    # ---------- Step 5：KMeans 聚类 ----------
    kmeans = KMeans(n_clusters=k, random_state=42)
    features["cluster"] = kmeans.fit_predict(X)

    # ---------- Step 6：二维可视化（avg_time × case_count） ----------
    plt.figure(figsize=(8, 6))
    for c in range(k):
        subset = features[features["cluster"] == c]
        plt.scatter(
            subset["avg_time"],
            subset["case_count"],
            label=f"Cluster {c}",
            alpha=0.7
        )

    plt.xlabel("Average Processing Time (hours)")
    plt.ylabel("Case Count")
    plt.title("Operator Behavior Clustering")
    plt.legend()
    plt.tight_layout()

    # ---------- Step 7：保存 ----------
    output_dir = Path("result/figures")
    output_dir.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_dir / "task3_2_operator_clustering.png", dpi=300)
    plt.close()

    # ---------- Step 8：保存聚类结果 ----------
    features.to_excel("result/result3.xlsx", index=False)

    return features


if __name__ == "__main__":
    df = preprocess_data("data/data.xlsx")

    print("Running Task 3.1...")
    analyze_processing_time_distribution(df)

    print("Running Task 3.2...")
    cluster_operator_behavior(df)
