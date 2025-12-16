# src/02_task1_statistics.py

import pandas as pd
import importlib.util
import os
from utils import calc_work_hours

# 动态导入 01_preprocess.py 模块
spec = importlib.util.spec_from_file_location("preprocess", os.path.join(os.path.dirname(__file__), "01_preprocess.py"))
preprocess_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(preprocess_module)
preprocess_data = preprocess_module.preprocess_data

df = preprocess_data("data/data.xlsx")

# 1. 只保留“完成的工序记录”
df_finished = df[df["is_finished"]].copy()

# 2. 找出“完成四道工序的案卷”

# 2.1 按案卷 + 工序去重
arch_flow = (
    df_finished
    .groupby(["sARCH_ID", "工序"])
    .size()
    .reset_index(name="cnt")
)

# 2.2 统计每个案卷完成了多少种工序
flow_count = (
    arch_flow
    .groupby("sARCH_ID")["工序"]
    .nunique()
    .reset_index(name="flow_num")
)

# 2.3 只保留完成 4 道工序的案卷
completed_archives = flow_count[flow_count["flow_num"] == 4]["sARCH_ID"]

# --- 任务 1.1 完成四道工序的案卷数量
num_completed_archives = completed_archives.nunique()

# 3. 汇总每个案卷 × 工序的开始 / 结束时间

# 3.1 只保留“完成四道工序”的案卷
df_valid = df_finished[df_finished["sARCH_ID"].isin(completed_archives)]


"""
这一步解决了：“同一案卷 + 同一工序多条记录怎么办？”
> 用 最早开始 + 最晚结束
"""
# 3.2 对每个案卷 × 工序汇总时间
flow_time_summary = (
    df_valid
    .groupby(["sARCH_ID", "工序"])
    .agg(
        start_time=("dUPDATE_TIME", "min"),
        end_time=("dNODE_TIME", "max")
    )
    .reset_index()
)

# 4. 把“长表”变成题目要求的“宽表（表 1）

# 4.1 透视成宽表
table = flow_time_summary.pivot(
    index="sARCH_ID",
    columns="工序",
    values=["start_time", "end_time"]
)

# 4.2 整理列名（为了导 Excel 和写报告）
table.columns = [
    f"{flow}_{t}"
    for t, flow in table.columns
]

table = table.reset_index()

# 5. 计算案卷完成时长（任务 1.1 的核心）

# 5.1 先算“工序级耗时”（只算 3 个工序）
valid_flows = ["扫描", "图像处理", "自检全检"]

flow_hours = (
    df_valid[df_valid["工序"].isin(valid_flows)]
    .groupby(["sARCH_ID", "工序"])["work_hours"]
    .sum()
    .reset_index()
)

# 5.2 汇总为“案卷完成时长”
archive_hours = (
    flow_hours
    .groupby("sARCH_ID")["work_hours"]
    .sum()
    .reset_index(name="完成时长")
)

archive_hours["完成时长"] = archive_hours["完成时长"].round(3)

"""
这一步结束，已经得到了 完整的表 1
"""
# 5.3 合并回表 1
result_table = table.merge(
    archive_hours,
    on="sARCH_ID",
    how="left"
)

# 6. 找出完成时长最长的 3 个案卷
top3 = (
    result_table
    .sort_values("完成时长", ascending=False)
    .head(3)
)


""""
在这一步已经完成了1.1
"""
# 7. 保存结果（题目明确要求）
result_table.to_excel(
    "result/result1_1.xlsx",
    index=False
)

"""
开始1.2
"""

# 1. 找出返工案卷（案卷级）
# 只看完成四道工序的案卷
df_completed = df[df["sARCH_ID"].isin(completed_archives)]

# 案卷级是否返工
rework_archives = (
    df_completed[df_completed["is_rework"]]
    ["sARCH_ID"]
    .unique()
)

num_rework_archives = len(rework_archives)

# 2. 计算返工案卷占比
rework_ratio = num_rework_archives / len(completed_archives) * 100
rework_ratio = round(rework_ratio, 3)

# ---- 构造题目要求的「表 2」 ----
# 3. 只保留返工案卷的记录
df_rework = df_completed[df_completed["sARCH_ID"].isin(rework_archives)]

# 4. 提取“返工工序 + 返工时间”

"""
返工时间 = dPROC_TIME
一个工序 最多只保留一次返工时间
如果有多次返工 → 取最早（稳妥）
"""
rework_summary = (
    df_rework[df_rework["is_rework"]]
    .groupby(["sARCH_ID", "工序"])
    .agg(
        rework_time=("dPROC_TIME", "min")
    )
    .reset_index()
)

# 5. 变成“案卷 × 工序”的宽表
table2 = rework_summary.pivot(
    index="sARCH_ID",
    columns="工序",
    values="rework_time"
).reset_index()

# 6. 保存 result1_2.xlsx
table2.to_excel(
    "result/result1_2.xlsx",
    index=False
)

"""
开始 1.3
"""

# 1. 只取「自检全检」工序的数据
df_check = df_completed[df_completed["工序"] == "自检全检"].copy()

# 2. 计算每个操作人员的“自检全检案卷总数”
total_archives = (
    df_check
    .groupby("iUSER_ID")["sARCH_ID"]
    .nunique()
    .reset_index(name="total_archives")
)

# 3. 计算每个操作人员的“返工案卷数”
rework_archives = (
    df_check[df_check["is_rework"]]
    .groupby("iUSER_ID")["sARCH_ID"]
    .nunique()
    .reset_index(name="rework_archives")
)

# 4. 合并并计算返工占比
result = total_archives.merge(
    rework_archives,
    on="iUSER_ID",
    how="left"
)

# 没有返工的人员，返工数为 0
result["rework_archives"] = result["rework_archives"].fillna(0)

result["返工案卷占比 (%)"] = (
    result["rework_archives"] / result["total_archives"] * 100
).round(3)

# 5. 排序并输出表 3
result_sorted = result.sort_values(
    "返工案卷占比 (%)",
    ascending=False
)

table3 = result_sorted[["iUSER_ID", "返工案卷占比 (%)"]]

table3.to_excel(
    "result/result1_3.xlsx",
    index=False
)

""""
开始 1.4
"""

# 1. 先只保留“完成记录”
df_f = df[df["is_finished"]].copy()

# 2. 统计每个工序完成案卷数量
archive_count = (
    df_f
    .groupby("工序")["sARCH_ID"]
    .nunique()
    .reset_index(name="完成案卷的数量")
)

# 3. 按「工序 + 批次」计算批次耗时（核心）

# 3.1 先聚合出批次时间区间
batch_time = (
    df_f
    .groupby(["工序", "sBatch_number"])
    .agg(
        batch_start=("dUPDATE_TIME", "min"),
        batch_end=("dNODE_TIME", "max")
    )
    .reset_index()
)

# 3.2 对每个批次计算“有效工作时长”
batch_time["batch_hours"] = batch_time.apply(
    lambda row: calc_work_hours(
        row["batch_start"],
        row["batch_end"]
    ),
    axis=1
)

# 4. 按工序汇总“总耗时”
total_hours = (
    batch_time
    .groupby("工序")["batch_hours"]
    .sum()
    .reset_index(name="总耗时 (h)")
)

total_hours["总耗时 (h)"] = total_hours["总耗时 (h)"].round(3)

# 5. 合并并计算平均耗时
result = archive_count.merge(
    total_hours,
    on="工序",
    how="left"
)

result["平均耗时 (h/卷)"] = (
    result["总耗时 (h)"] / result["完成案卷的数量"]
).round(3)

# 6. 整理并保存 result1_4.xlsx
table4 = result[[
    "工序",
    "完成案卷的数量",
    "总耗时 (h)",
    "平均耗时 (h/卷)"
]]

table4.to_excel(
    "result/result1_4.xlsx",
    index=False
)
