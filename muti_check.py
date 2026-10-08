import csv
import json
from pathlib import Path


清洗清单路径 = Path("药品名规格清洗清单.csv")
原始药品库路径 = Path("原药品库.csv")
输出路径 = Path("药品归并核对.json")

with 清洗清单路径.open("r", encoding="utf-8-sig", newline="") as 文件:
    清洗记录 = list(csv.DictReader(文件))

with 原始药品库路径.open("r", encoding="utf-8-sig", newline="") as 文件:
    原始记录 = list(csv.DictReader(文件))

原始记录索引 = {记录["id"]: 记录 for 记录 in 原始记录}
归并核对清单 = []

for 清洗项 in 清洗记录:
    if int(清洗项["来源记录数"]) <= 1:
        continue

    来源药品ID = 清洗项["来源药品ID"].split("|")
    归并核对清单.append(
        {
            "药品ID": 清洗项["药品ID"],
            "原始药品记录": [原始记录索引[药品ID] for 药品ID in 来源药品ID],
        }
    )

结果 = {
    "说明": "按归并组列出清洗清单保留的药品ID及对应原始药品库记录，供核对是否可以归并。",
    "归并组数": len(归并核对清单),
    "归并来源记录数": sum(len(项目["原始药品记录"]) for 项目 in 归并核对清单),
    "归并核对清单": 归并核对清单,
}

with 输出路径.open("w", encoding="utf-8") as 文件:
    json.dump(结果, 文件, ensure_ascii=False, indent=2)
    文件.write("\n")

print(f"已输出 {输出路径}，共 {结果['归并组数']} 组、{结果['归并来源记录数']} 条原始记录。")
