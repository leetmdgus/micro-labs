import win32com.client as win32
import pythoncom
import os
import uuid

TEMPLATE_DIR = "./templates"
OUTPUT_DIR = "./outputs"

def generate_hwp(template_name: str, fields: dict):
    pythoncom.CoInitialize()  # COM 초기화
    hwp = None
    output_file = os.path.join(OUTPUT_DIR, f"{uuid.uuid4()}.hwp")

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        # HWP 객체 생성 직후 추가 보안 모듈 해제 (코드에서 처리)
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")

        template_path = os.path.join(TEMPLATE_DIR, template_name)
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"템플릿 없음: {template_path}")

        if not hwp.Open(template_path):
            raise RuntimeError("HWP 템플릿 열기 실패")

        # 필드 채우기
        for key, value in fields.items():
            try:
                hwp.PutFieldText(key, value)
            except Exception as e:
                print(f"⚠️ 필드 {key} 채우기 실패: {e}")

        hwp.SaveAs(output_file)
        print(f"✅ 생성 완료: {output_file}")
        return output_file
    finally:
        if hwp:
            try:
                hwp.Quit()
            except:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    generate_hwp(
        template_name="2025년도_예비창업패키지_사업계획서_양식.hwp",
        fields={
            "INPUT_FIELD": "히히",
            "PROJECT_NAME": "단택시",
            "1_PROBLEM": "고령자의 택시 호출 불편"
        }
    )
