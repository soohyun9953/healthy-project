import os
import sys
import tempfile
from flask import Flask, request, jsonify, send_file

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

    # PowerPoint OLE 엔진 초기화 및 변환 (백그라운드 실행)
    powerpoint = None
    presentation = None
    try:
        # 멀티스레드 환경 대비 CoInitialize 실행
        import pythoncom
        pythoncom.CoInitialize()

        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # win32com PPT 변환 시 Visible=True 상태여야 폰트/표 스타일 레이아웃이 원본과 100% 동일하게 완벽 렌더링됨
        powerpoint.Visible = True 
        
        # WithWindow=False를 제거하여 OLE 창 충돌로 인한 렌더링 에러를 완전히 우회합니다.
        presentation = powerpoint.Presentations.Open(abs_input_path)
        # 32는 ppSaveAsPDF 상수값
        presentation.SaveAs(output_pdf_path, 32)
        presentation.Close()
        
        return True, output_pdf_path
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
        
    # 임시 폴더에 파일 저장 후 변환
    temp_dir = tempfile.gettempdir()
    temp_input_path = os.path.join(temp_dir, uploaded_file.filename)
    uploaded_file.save(temp_input_path)
    
    success, result_msg = convert_ppt_to_pdf_local(temp_input_path)
    
    # 원본 임시 파일 즉시 청소
    if os.path.exists(temp_input_path):
        try: os.remove(temp_input_path)
        except: pass
        
    if success:
        pdf_path = result_msg
        # 변환된 PDF를 브라우저로 직접 전송
        return send_file(pdf_path, as_attachment=True, download_name=os.path.basename(pdf_path))
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
