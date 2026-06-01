import pandas as pd
import sys

# Set encoding for output to UTF-8
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def analyze_file(filename):
    print(f"Analyzing {filename}...")
    try:
        xls = pd.ExcelFile(filename)
        for sheet_name in xls.sheet_names:
            print(f"\n--- Sheet: {sheet_name} ---")
            df = pd.read_excel(filename, sheet_name=sheet_name)
            print(df.head())
            
            # Search for "커뮤니케이션" or "협업"
            mask = df.stack().astype(str).str.contains('커뮤니케이션|협업').any()
            if mask:
                print(f"Found relevant terms in sheet: {sheet_name}")
                # Filter rows
                relevant_rows = df[df.apply(lambda row: row.astype(str).str.contains('커뮤니케이션|협업').any(), axis=1)]
                print(relevant_rows)
    except Exception as e:
        print(f"Error: {e}")

analyze_file('데이터_아키텍처_정의_산출물_점검_결과_전수.xlsx')
analyze_file('데이터_아키텍처_정의_산출물_점검_결과.xlsx')
