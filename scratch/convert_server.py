import os
import sys
import tempfile
import shutil
import uuid
from flask import Flask, request, jsonify, send_file, after_this_request

# 플라스크 인스턴스 생성 (한글 설명 및 스네이크 케이스 준수)
app = Flask(__name__)

# CORS 직접 구현 (외부 flask_cors 패키지 없이도 브라우저 CORS 차단 완벽 우회)
@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
    return response

# 헬퍼 함수: win32com을 활용해 로컬 PPTX 파일을 PDF로 변환
def convert_ppt_to_pdf_local(input_file_path):
    try:
        import win32com.client
    except ImportError:
        return False, "pywin32 라이브러리가 로컬 파이썬에 설치되어 있지 않습니다. 터미널에서 'pip install pywin32'를 구동한 후 서버를 다시 켜주세요."

    if not os.path.exists(input_file_path):
        return False, f"입력 파일을 찾을 수 없습니다: {input_file_path}"

    # 절대 경로로 안전하게 정규화
    abs_input_path = os.path.abspath(input_file_path)
    base_name = os.path.splitext(abs_input_path)[0]
    output_pdf_path = base_name + ".pdf"

    # 💡 [유니코드 경로 크래시 완전 해결]
    # 파일명에 한글, 공백, 유니코드 특수기호(Ⅳ 등)가 섞여 있으면 PowerPoint COM OLE 엔진이 경로 연산에 폭사합니다.
    # 이를 원천 방지하기 위해 파일 사본을 시스템 Temp 폴더에 단순 영문명(UUID)으로 임시 복제하여 변환 후 최종 이동시킵니다.
    temp_dir = tempfile.gettempdir()
    unique_id = uuid.uuid4().hex
    safe_temp_pptx = os.path.join(temp_dir, f"ole_in_{unique_id}.pptx")
    safe_temp_pdf = os.path.join(temp_dir, f"ole_out_{unique_id}.pdf")

    try:
        shutil.copy2(abs_input_path, safe_temp_pptx)
    except Exception as e:
        return False, f"임시 유니코드 안전 사본 복제 에러: {str(e)}"

    powerpoint = None
    presentation = None
    try:
        # 멀티스레드 환경 대비 CoInitialize 실행
        import pythoncom
        pythoncom.CoInitialize()

        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # win32com PPT 변환 시 Visible=True 상태여야 폰트/표 스타일 레이아웃이 원본과 100% 동일하게 완벽 렌더링됨
        powerpoint.Visible = True 
        
        # 유니코드 안전 사본 경로로 OLE 로드
        presentation = powerpoint.Presentations.Open(safe_temp_pptx)
        # 32는 ppSaveAsPDF 상수값
        presentation.SaveAs(safe_temp_pdf, 32)
        presentation.Close()
        
        # 생성된 임시 PDF를 원래 타겟 경로로 이송
        if os.path.exists(safe_temp_pdf):
            shutil.copy2(safe_temp_pdf, output_pdf_path)
            return True, output_pdf_path
        else:
            return False, "PDF 변환 파일이 생성되지 않았습니다."
    except Exception as e:
        return False, f"PowerPoint 렌더러 변환 에러: {str(e)}"
    finally:
        if presentation:
            try: presentation.Close()
            except: pass
        if powerpoint:
            try: powerpoint.Quit()
            except: pass
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except:
            pass
        # 임시 사본 청소
        for temp_file in [safe_temp_pptx, safe_temp_pdf]:
            if os.path.exists(temp_file):
                try: os.remove(temp_file)
                except: pass

# API 1: 서버 구동 여부 확인 (CORS 연동용)
@app.route("/status", methods=["GET", "OPTIONS"])
def check_server_status():
    if request.method == "OPTIONS":
        return "", 200
    return jsonify({
        "status": "online",
        "message": "로컬 PPT PDF 변환 헬퍼 서버가 원활하게 작동 중입니다."
    }), 200

# API 2: 로컬 절대 경로를 통한 변환 및 동일 디렉토리 저장
@app.route("/convert-local", methods=["POST", "OPTIONS"])
def handle_local_conversion():
    if request.method == "OPTIONS":
        return "", 200
        
    request_data = request.get_json() or {}
    local_path = request_data.get("local_file_path", "").strip()
    
    if not local_path:
        return jsonify({"success": False, "message": "로컬 파일 절대 경로가 누락되었습니다."}), 400
        
    success, result_msg = convert_ppt_to_pdf_local(local_path)
    
    if success:
        return jsonify({
            "success": True,
            "message": f"동일 디렉토리에 PDF 파일이 성공적으로 생성되었습니다!",
            "pdf_path": result_msg
        }), 200
    else:
        return jsonify({
            "success": False,
            "message": result_msg
        }), 500

# API 3: 브라우저 드래그앤드롭 업로드 변환 후 바이너리로 직접 응답
@app.route("/convert-upload", methods=["POST", "OPTIONS"])
def handle_upload_conversion():
    if request.method == "OPTIONS":
        return "", 200
        
    if "file" not in request.files:
        return jsonify({"success": False, "message": "업로드된 파일이 없습니다."}), 400
        
    uploaded_file = request.files["file"]
    if uploaded_file.filename == "":
        return jsonify({"success": False, "message": "파일명이 유효하지 않습니다."}), 400
        
    # 💡 [유니코드 뇌관 완전 해결] 처음 임시 파일 저장 시 한글/특수문자가 OLE 파싱 경로에 엉키지 않도록 완전한 ASCII로 임시 세이브
    temp_dir = tempfile.gettempdir()
    unique_id = uuid.uuid4().hex
    temp_input_path = os.path.join(temp_dir, f"upload_in_{unique_id}.pptx")
    uploaded_file.save(temp_input_path)
    
    success, result_msg = convert_ppt_to_pdf_local(temp_input_path)
    
    # 원본 임시 파일 즉시 청소
    if os.path.exists(temp_input_path):
        try: os.remove(temp_input_path)
        except: pass
        
    if success:
        pdf_path = result_msg
        
        # 원래 업로드한 파일명의 확장자를 .pdf로 매핑하여 안전한 다운로드 제공
        original_base = os.path.splitext(uploaded_file.filename)[0]
        download_pdf_name = f"{original_base}.pdf"
        
        # 💡 [리소스 무결성 청소] 응답 전송이 완료된 후 백그라운드에서 임시 생성되었던 PDF를 소거해 디스크 오염 원천 방어
        @after_this_request
        def remove_temporary_pdf(response):
            try:
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
            except Exception as e:
                app.logger.error(f"Error removing temp pdf file: {e}")
            return response

        return send_file(pdf_path, as_attachment=True, download_name=download_pdf_name)
    else:
        return jsonify({
            "success": False,
            "message": result_msg
        }), 500

if __name__ == "__main__":
    print("------------------------------------------------------------------")
    print("   로컬 PPT -> PDF 변환 헬퍼 서버가 포트 5000에서 구동을 시작합니다.   ")
    print("   이 창을 열어둔 상태에서 웹 사이트의 변환 기능을 사용해 주세요.    ")
    print("------------------------------------------------------------------")
    app.run(host="127.0.0.1", port=5000, debug=False)
