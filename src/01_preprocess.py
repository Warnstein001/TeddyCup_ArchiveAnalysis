# src/01_preprocess.py

import pandas as pd
from utils import calc_work_hours

def preprocess_data(data_path: str) -> pd.DataFrame:
    """"
    读取并预处理原始数据，返回可分析的 DataFrame
    """
    # 1. 读取数据
    df = pd.read_excel(data_path)
    
    # 2. 时间字段统一转换
    time_cols = ["dUPDATE_TIME", "dNODE_TIME", "dPROC_TIME"]
    for col in time_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    
    # 3. 工序编号 -> 中文名称
    flow_map = {
        1: "扫描",
        2: "图像处理",
        3: "自检全检",
        4: "PDF处理"
    }
    df["工序"] = df["iFLOW_NODE_NO"].map(flow_map)

    # 4. 计算每条记录的工序有效时长
    df["work_hours"] = df.apply(
        lambda row: calc_work_hours(
            row["dUPDATE_TIME"],
            row["dNODE_TIME"]
        ),
        axis=1
    )

    # 5. 常用分析辅助字段
    df["finish_date"] = df["dNODE_TIME"].dt.date
    df["is_rework"] = df["iNODE_STATUS"] == 5
    df["is_finished"] = df["iNODE_STATUS"].isin([2, 5])

    return df

if __name__ == "__main__":
     # 手动运行时用于检查
    df_clean = preprocess_data("data/data.xlsx")
    
    # 创建输出文件路径
    output_file = "result/preprocessed_data_analysis.xlsx"
    
    # 使用ExcelWriter创建多个工作表
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 1. 完整的预处理数据
        df_clean.to_excel(writer, sheet_name='预处理数据', index=False)
        
        # 2. 工序分布统计
        flow_stats = df_clean["工序"].value_counts().reset_index()
        flow_stats.columns = ['工序', '数量']
        flow_stats.to_excel(writer, sheet_name='工序分布', index=False)
        
        # 3. 工作时长统计
        work_hours_stats = df_clean["work_hours"].describe().reset_index()
        work_hours_stats.columns = ['统计项', '数值']
        work_hours_stats.to_excel(writer, sheet_name='工作时长统计', index=False)
        
        # 4. 状态分布
        status_stats = df_clean["iNODE_STATUS"].value_counts().reset_index()
        status_stats.columns = ['状态码', '数量']
        status_stats.to_excel(writer, sheet_name='状态分布', index=False)
        
        # 5. 返工和完成情况
        summary_stats = pd.DataFrame({
            '项目': ['返工记录', '非返工记录', '已完成记录', '未完成记录'],
            '数量': [
                df_clean["is_rework"].sum(),
                (~df_clean["is_rework"]).sum(),
                df_clean["is_finished"].sum(),
                (~df_clean["is_finished"]).sum()
            ]
        })
        summary_stats.to_excel(writer, sheet_name='汇总统计', index=False)
        
        # 6. 按工序的详细统计
        flow_detail = df_clean.groupby('工序').agg({
            'work_hours': ['count', 'mean', 'std', 'min', 'max'],
            'is_rework': 'sum',
            'is_finished': 'sum'
        }).round(4)
        flow_detail.columns = ['记录数', '平均工时', '工时标准差', '最小工时', '最大工时', '返工数', '完成数']
        flow_detail.to_excel(writer, sheet_name='按工序统计')
    
    print(f"数据已成功输出到: {output_file}")
    print(f"包含以下工作表:")
    print("- 预处理数据: 完整的处理后数据")
    print("- 工序分布: 各工序的记录数量")
    print("- 工作时长统计: 工时的描述性统计")
    print("- 状态分布: 各状态码的分布")
    print("- 汇总统计: 返工和完成情况汇总")
    print("- 按工序统计: 按工序分组的详细统计")
    
    # 同时在控制台显示基本信息
    print(f"\n数据基本信息: 共{df_clean.shape[0]}行, {df_clean.shape[1]}列")