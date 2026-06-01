import requests
import os
import json

# 테스트 대상 파일 설정
test_file = r"c:\Users\JEJU_MEC_3\Desktop\AI Coding\3.2.1-6 정보시스템 데이터 아키텍처 정의 _v0.7(D).pptx"
api_url = "http://127.0.0.1:5000/convert-local"

print(f"테스트 파일 경로: {test_file}")
print(f"파일 존재 여부: {os.path.exists(test_file)}")

if os.path.exists(test_file):
    payload = {"local_file_path": test_file}
    headers = {"Content-Type": "application/json"}
    
    print("PDF 변환 요청 전송 중...")
    try:
        response = requests.post(api_url, data=json.dumps(payload), headers=headers, timeout=60)
        print(f"응답 코드: {response.status_code}")
        response_data = response.json()
        print(f"응답 본문: {json.dumps(response_data, indent=4, ensure_ascii=False)}")
        
        pdf_path = response_data.get("pdf_path")
        if pdf_path and os.path.exists(pdf_path):
            print(f"성공! PDF 파일이 동일 디렉토리에 정상적으로 생성되었습니다: {pdf_path}")
        else:
            print("오류: PDF 파일이 디스크에 생성되지 않았습니다.")
    except Exception as e:
        print(f"요청 중 에러 발생: {e}")
else:
    print("오류: 테스트용 PPTX 파일이 해당 경로에 존재하지 않습니다.")
