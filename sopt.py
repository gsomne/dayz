import pandas as pd
from collections import defaultdict
import csv

with open("huita.csv", "r", encoding="utf-8") as input_file:
    reader = csv.reader(input_file)
    next(reader)
    class_data = []

    for row in reader:
        if not row or not row[0].strip():
            continue
        if row[0][-1] == "»":
            row[0] = row[0].replace("«", "").replace("»", "")
        class_data.append(row)

grouped_classes = defaultdict(lambda: defaultdict(list))

for entry in class_data:
    key_name = entry[0].strip()
    key_find = key_name.split()[-1]
    class_name = entry[1].strip() if len(entry) > 1 else ""
    if key_find and class_name:
        grouped_classes[key_find][key_name].append(class_name)

for group_key in list(grouped_classes.keys()):
    grouped_classes[group_key] = {
        k: v for k, v in grouped_classes[group_key].items() if any(s.strip() for s in v)
    }
    if not grouped_classes[group_key]:
        del grouped_classes[group_key]

with pd.ExcelWriter("grouped_sorted.xlsx", engine="openpyxl") as writer:
    for group, items in grouped_classes.items():
        data = []
        for full_name, class_list in items.items():
            for class_name in class_list:
                data.append([group, full_name, class_name])
        df = pd.DataFrame(data, columns=["Group", "Full Name", "Class Name"])
        df.to_excel(writer, sheet_name=group[:31], index=False)
