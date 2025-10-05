# 🧪 Week01a — Template Merge (DOCX / TXT / MD 자동 치환)

> CSV 데이터를 불러와 DOCX / TXT / MD 템플릿의 자리표시자를 자동으로 채워주는 문서 자동화 실험  
> *(DocxTemplate + Pandas + Jinja2 기반)*

## 📘 개요  
이 프로젝트는 **문서 자동 생성 자동화 실험**입니다.  
`data.csv`의 각 행(row)을 읽어와,  
`data/template.docx`, `data/template.txt` 파일의 필드(`{{필드명}}`)에 자동으로 데이터를 삽입하여  
각각 `outputs_docx/`, `outputs_md/`, `outputs_txt/` 폴더에 개별 `.docx`, `.md`, `.txt` 파일로 저장합니다.

## ⚙️ 기술 스택  
- **언어:** Python 3.13
- **라이브러리:** pandas, docxtpl, os, jinja2
- **환경:** venv (가상환경)  

## 📁 폴더 구조
week01a_template_merge/
┣ data/ # CSV 데이터 및 템플릿 파일 (DOCX / MD / TXT)
┣ outputs_docx/ # DOCX 결과 파일 (자동 생성됨)
┣ outputs_md/ # MD 결과 파일 (자동 생성됨)
┣ outputs_txt/ # TXT 결과 파일 (자동 생성됨)
┣ main_docx.py # DOCX 병합 코드
┣ main_md.py # MD 병합 코드
┣ main_txt.py # TXT 병합 코드
┣ requirements.txt 
┗ README.md

## 실행방법
### 1️⃣ 가상환경 생성 및 실행
```bash
python -m venv venv
source venv/Scripts/activate   # windows bash
venv\Scripts\activate.bat # windows cmd
venv\Scripts\Activate.ps1 # windows powershell
# 또는
source venv/bin/activate       # macOS / Linux
```

### 2️⃣ 라이브러리 설치
```bash
pip install -r requirements.txt
```

### 3️⃣ 실행
```bash
python main_docx.py   # DOCX 문서 자동 생성
python main_md.py     # Markdown 파일 자동 생성
python main_txt.py    # Text 파일 자동 생성
```

## 📊 결과 예시

실행 후 각 폴더에 다음과 같은 결과가 생성됩니다:

outputs_docx/
 ┣ output_1.docx
 ┣ output_2.docx
 ┗ output_3.docx