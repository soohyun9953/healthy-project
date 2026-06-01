import os
from playwright.sync_api import sync_playwright

url = 'http://localhost:5173'
ppt_path = r'c:\Users\JEJU_MEC_3\Desktop\AI Coding\scratch\test.pptx'

if not os.path.exists(ppt_path):
    print(f"ERROR: PPT file not found at {ppt_path}")
    exit(1)

console_logs = []

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(viewport={'width': 1920, 'height': 1080})

    def handle_console_message(msg):
        log_line = f"[{msg.type.upper()}] {msg.text}"
        console_logs.append(log_line)
        print(f"Browser Console: {log_line}")

    page.on("console", handle_console_message)

    # 1. 페이지 접속
    print(f"Navigating to {url}...")
    page.goto(url)
    page.wait_for_load_state('networkidle')
    page.wait_for_timeout(2000)

    # 2. 사이드바 PPT 생성(표준산출물) 메뉴 클릭
    print("Clicking sidebar tab for PPT generator...")
    page.locator("#sidebar-tab-ppt").click()
    page.wait_for_timeout(1000)

    # 3. 일괄 편집 탭 클릭
    print("Switching to PPT 일괄 편집 tab...")
    page.locator("#tab-batch-edit").click()
    page.wait_for_timeout(1000)

    # 4. PPT 파일 업로드
    print(f"Uploading file: {ppt_path}")
    file_input = page.locator("input[id='batch-ppt-upload']")
    file_input.set_input_files(ppt_path)
    page.wait_for_timeout(5000) # 대용량 파일 로딩 대기

    # 5. 옵션 F 체크박스 클릭
    print("Clicking Option F checkbox...")
    page.locator("#checkbox-option-f").click()
    page.wait_for_timeout(1000)

    # 5-2. 옵션 G 체크박스 클릭
    print("Clicking Option G checkbox...")
    page.locator("#checkbox-option-g").click()
    page.wait_for_timeout(1000)

    # 6. 일괄 편집 실행 버튼 클릭
    print("Clicking process button...")
    page.on("dialog", lambda dialog: dialog.accept())
    
    try:
        # showSaveFilePicker가 뜰 때 Playwright에서 오류가 발생할 수 있으므로 에러 예외처리 및 확실하게 클릭
        page.locator("#btn-batch-process").click()
        print("Waiting for batch processing...")
        page.wait_for_timeout(25000) # 처리 시간을 25초 정도 넉넉히 대기
    except Exception as e:
        print(f"Click or wait failed: {e}")

    # 스크린샷 캡처
    screenshot_path = r'c:\Users\JEJU_MEC_3\Desktop\AI Coding\scratch\result_screenshot.png'
    page.screenshot(path=screenshot_path)
    print(f"Screenshot saved to {screenshot_path}")

    browser.close()

# 로그 결과 출력
print("\n--- FINAL CONSOLE LOGS ---")
for log in console_logs:
    if "error" in log.lower() or "exception" in log.lower() or "failed" in log.lower():
        print(f"\033[91m{log}\033[0m")
    else:
        print(log)
