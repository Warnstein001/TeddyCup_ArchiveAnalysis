# src/utils.py
import pandas as pd
from datetime import datetime, time, timedelta


# 工作时间定义
WORK_START_MORNING = time(8, 30)
WORK_END_MORNING = time(12, 0)
WORK_START_AFTERNOON = time(13, 0)
WORK_END_AFTERNOON = time(18, 0)

def is_workday(dt: datetime) -> bool:
    """是否为周一 ~ 周六"""
    return dt.weekday() < 6

def _calc_overlap(st1, ed1, st2, ed2):
    """"
    计算两个时间区间重叠的秒数
    """
    latest_st = max(st1, st2)
    earliest_ed = min(ed1, ed2)
    delta = (earliest_ed - latest_st).total_seconds()
    return max(0, delta)

def calc_work_hours(st: datetime, ed: datetime) -> float:
    """"
    计算 start ~ end 之间的有效工作时长（单位：h）
    """
    if pd.isna(st) or pd.isna(ed):
        return 0.0
    if st >= ed:
        return 0.0
    total_seconds = 0
    cur_date = st.date()
    end_date = ed.date()
    while cur_date <= end_date:
        day_start = datetime.combine(cur_date, time(0, 0))
        day_end = datetime.combine(cur_date, time(23, 59, 59))

        # 该日实际参与计算的时间区间
        interval_start = max(st, day_start)
        interval_end = min(ed, day_end)

        if interval_start < interval_end and is_workday(interval_start):
             # 上午
            morning_start = datetime.combine(cur_date, WORK_START_MORNING)
            morning_end = datetime.combine(cur_date, WORK_END_MORNING)

            # 下午
            afternoon_start = datetime.combine(cur_date, WORK_START_AFTERNOON)
            afternoon_end = datetime.combine(cur_date, WORK_END_AFTERNOON)

            total_seconds += _calc_overlap(
                interval_start, interval_end,
                morning_start, morning_end
            )

            total_seconds += _calc_overlap(
                interval_start, interval_end,
                afternoon_start, afternoon_end
            )

        cur_date += timedelta(days=1)

    return round(total_seconds / 3600, 6)
