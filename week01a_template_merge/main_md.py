import pandas as pd
from jinja2 import Template
import os

# 1) 데이터 로드
df = pd.read_csv("./data/data.csv")

# 2) 템플릿 읽기
with open("./data/template.txt", "r", encoding="utf-8") as f:
    template_str = f.read()

template = Template(template_str)

# 3) 각 row별로 치환 수행
# 저장 폴더
output_dir = "./outputs_md"
os.makedirs(output_dir, exist_ok=True)  # 폴더 없으면 자동 생성

for idx, row in df.iterrows():
    output = template.render(**row.to_dict())
    filename = os.path.join(output_dir, f"output_{idx+1}.md")
    with open(filename, "w", encoding="utf-8") as f:
        f.write(output)
    print(f"✅ 생성 완료: {filename}")