from datetime import datetime
import sys
import os

# 添加父目录到路径，这样可以导入 utils
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from utils import calc_work_hours

# 情况 1：同一天、跨中午
s = datetime(2020, 7, 14, 10, 0)
e = datetime(2020, 7, 14, 15, 0)
res = calc_work_hours(s, e)
if res == 4:
    print(res, "✔")
else:
    print("❌")
# 10:00-12:00 + 13:00-15:00 = 4h

# 情况 2：跨天
s = datetime(2020, 7, 14, 17, 0)
e = datetime(2020, 7, 15, 9, 0)
res = calc_work_hours(s, e)
if res == 1.5:
    print(res, "✔")
else:
    print("❌")
# 17:00-18:00 + 8:30-9:00 = 1.5h

# 情况 3：周日
s = datetime(2020, 7, 12, 10, 0)  # 周日
e = datetime(2020, 7, 12, 16, 0)
res = calc_work_hours(s, e)
if res == 0:
    print(res, "✔")
else:
    print("❌")
# 应为 0
