import pandas as pd
from docxtpl import DocxTemplate
import os

# 1) 데이터 로드
df = pd.read_csv("./data/data.csv", encoding="utf-8-sig")

# 2) 템플릿 불러오기
template_path = "./data/template.docx"

# 3) 출력 폴더 생성
output_dir = "./outputs_docx"
os.makedirs(output_dir, exist_ok=True)

# 4) 각 row별 치환 수행
for idx, row in df.iterrows():
    doc = DocxTemplate(template_path)
    context = row.to_dict()  # CSV 컬럼 → 템플릿 키 매핑
    doc.render(context)
    
    filename = os.path.join(output_dir, f"output_{idx+1}.docx")
    doc.save(filename)
    print(f"✅ 생성 완료: {filename}")