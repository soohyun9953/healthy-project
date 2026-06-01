import zipfile
import re
import sys
import os
import shutil
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')

pptx_path = r"C:\Users\KITC\Downloads\바이오파운드리 ISMP_Ⅳ.5. 정보시스템 요건 기술서.pptx"
hwpx_path = r"C:\Users\KITC\Downloads\바이오파운드리 통합플랫폼 구축 사업 제안요청서_v0.1.hwpx"
output_path = r"C:\Users\KITC\Downloads\융합_바이오파운드리 통합플랫폼 구축 사업 제안요청서_v0.1.hwpx"

# XML 네임스페이스 원본 접두사 강제 등록 (한글 규격 호환성 보장)
ET.register_namespace('hp', 'http://www.hancom.co.kr/hwpml/2011/paragraph')
ET.register_namespace('hs', 'http://www.hancom.co.kr/hwpml/2011/section')
ET.register_namespace('ha', 'http://www.hancom.co.kr/hwpml/2011/app')
ET.register_namespace('hc', 'http://www.hancom.co.kr/hwpml/2011/core')
ET.register_namespace('hh', 'http://www.hancom.co.kr/hwpml/2011/head')
ET.register_namespace('hm', 'http://www.hancom.co.kr/hwpml/2011/master-page')
ET.register_namespace('hpf', 'http://www.hancom.co.kr/schema/2011/hpf')
ET.register_namespace('config', 'urn:oasis:names:tc:opendocument:xmlns:config:1.0')
ET.register_namespace('hp10', 'http://www.hancom.co.kr/hwpml/2016/paragraph')
ET.register_namespace('hhs', 'http://www.hancom.co.kr/hwpml/2011/history')

class HwpxCharPrRegistry:
    def __init__(self, header_xml_path):
        self.header_xml_path = header_xml_path
        self.tree = ET.parse(header_xml_path)
        self.root = self.tree.getroot()
        self.ns = {
            'hh': 'http://www.hancom.co.kr/hwpml/2011/head'
        }
        ET.register_namespace('hh', 'http://www.hancom.co.kr/hwpml/2011/head')
        
        self.char_properties = self.root.find('.//hh:charProperties', self.ns)
        if self.char_properties is None:
            self.char_properties = self.root.find('.//charProperties')
            
        self.registry = {}  # key: (base_id, bold, color) -> charPr_id
        self.max_id = -1
        
        if self.char_properties is not None:
            for charPr in self.char_properties:
                c_id = int(charPr.get('id', -1))
                if c_id > self.max_id:
                    self.max_id = c_id

    def get_or_create_char_pr(self, base_char_pr_id, bold, color):
        """
        base_char_pr_id: 오리지널 charPr ID (예: '10')
        bold: bool (굵게 속성 여부)
        color: str (예: '#FF0000', 16진수 RGB 컬러 코드)
        """
        base_pr = None
        for pr in self.char_properties:
            if pr.get('id') == str(base_char_pr_id):
                base_pr = pr
                break
                
        if base_pr is None:
            for pr in self.char_properties:
                if pr.get('id') == '0':
                    base_pr = pr
                    break
                    
        if base_pr is None:
            base_pr = self.char_properties[0]
            
        target_color = color if color is not None else base_pr.get('textColor', '#000000')
        
        cache_key = (str(base_char_pr_id), bold, target_color)
        if cache_key in self.registry:
            return self.registry[cache_key]
            
        # 부모 스타일과 아예 동일하면 원래 ID 반환
        parent_has_bold = base_pr.find('{http://www.hancom.co.kr/hwpml/2011/head}bold') is not None or base_pr.find('bold') is not None
        parent_color = base_pr.get('textColor', '#000000')
        
        if parent_has_bold == bold and parent_color == target_color:
            return base_char_pr_id
            
        # 새 글자 모양 복제 및 추가
        self.max_id += 1
        new_id = str(self.max_id)
        
        new_pr = ET.fromstring(ET.tostring(base_pr))
        new_pr.set('id', new_id)
        new_pr.set('textColor', target_color)
        
        bold_tag_name = '{http://www.hancom.co.kr/hwpml/2011/head}bold'
        bold_elem = new_pr.find(bold_tag_name)
        if bold_elem is None:
            bold_elem = new_pr.find('bold')
            
        if bold:
            if bold_elem is None:
                new_bold = ET.Element(bold_tag_name)
                inserted = False
                for idx, child in enumerate(list(new_pr)):
                    tag_name = child.tag.split('}')[-1]
                    if tag_name in ['underline', 'strikeout', 'outline', 'shadow']:
                        new_pr.insert(idx, new_bold)
                        inserted = True
                        break
                if not inserted:
                    new_pr.append(new_bold)
        else:
            if bold_elem is not None:
                new_pr.remove(bold_elem)
                
        self.char_properties.append(new_pr)
        self.char_properties.set('itemCnt', str(len(self.char_properties)))
        
        self.registry[cache_key] = new_id
        return new_id

    def save(self):
        self.tree.write(self.header_xml_path, encoding='utf-8', xml_declaration=True)

def extract_cell_styled_text(tc, ns):
    """
    a:tc 내부의 단락 및 런들을 순회하며 서식 정보(굵게, 색상)가 포함된 텍스트 구조를 가져옵니다.
    """
    paragraphs = []
    txBody = tc.find('a:txBody', ns)
    if txBody is None:
        return paragraphs
        
    for p in txBody.findall('a:p', ns):
        p_runs = []
        for child in p:
            tag = child.tag.split('}')[-1]
            if tag == 'r':
                text_elem = child.find('a:t', ns)
                if text_elem is not None and text_elem.text:
                    bold = False
                    color = None
                    
                    rPr = child.find('a:rPr', ns)
                    if rPr is not None:
                        if rPr.attrib.get('b') in ['1', 'true']:
                            bold = True
                        solidFill = rPr.find('a:solidFill', ns)
                        if solidFill is not None:
                            srgbClr = solidFill.find('a:srgbClr', ns)
                            if srgbClr is not None and 'val' in srgbClr.attrib:
                                color = '#' + srgbClr.attrib['val']
                                
                    p_runs.append({
                        'text': text_elem.text,
                        'bold': bold,
                        'color': color
                    })
            elif tag == 'br':
                p_runs.append({
                    'text': '\n',
                    'bold': False,
                    'color': None
                })
        paragraphs.append(p_runs)
    return paragraphs

def get_plain_text(paras):
    """
    런 구조의 문단 리스트를 단순 줄글로 취합해 줍니다 (유효성 검사 및 백업용).
    """
    lines = []
    for para in paras:
        line = "".join([run['text'] for run in para])
        lines.append(line)
    return "\n".join(lines).strip()

def parse_pptx_requirements():
    print("1. PPTX 파일 파싱 중...")
    with zipfile.ZipFile(pptx_path, 'r') as z:
        slides = sorted([f for f in z.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')],
                        key=lambda x: int(re.search(r'\d+', x).group()))
        
        ns = {
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        }
        
        requirements = []
        id_pattern = re.compile(r'^[A-Za-z0-9]+[\-_][A-Za-z0-9]+$')
        
        for slide_path in slides:
            xml_data = z.read(slide_path)
            root = ET.fromstring(xml_data)
            
            tbls = root.findall('.//a:tbl', ns)
            for tbl in tbls:
                rows = tbl.findall('a:tr', ns)
                if not rows:
                    continue
                
                # 1단계: 첫 번째 행(헤더 행)을 검사하여 각 매핑 열의 실제 인덱스 감지
                first_row_cells = rows[0].findall('a:tc', ns)
                header_texts = []
                for tc in first_row_cells:
                    h_text = "".join([tx.text for tx in tc.findall('.//a:t', ns) if tx.text]).strip()
                    header_texts.append(h_text)
                
                id_idx, name_idx, desc_idx, detail_idx = 0, 1, 2, 3
                for c_idx, h_text in enumerate(header_texts):
                    if "고유번호" in h_text or "ID" in h_text:
                        id_idx = c_idx
                    elif "명칭" in h_text or "요구사항명" in h_text:
                        name_idx = c_idx
                    elif "정의" in h_text or "개요" in h_text:
                        desc_idx = c_idx
                    elif "세부내용" in h_text or "요건" in h_text or "상세설명" in h_text:
                        detail_idx = c_idx
                
                # 2단계: 행 순회 가동 및 매핑 데이터 수집 (스타일 구조 보존)
                for tr in rows:
                    cells = tr.findall('a:tc', ns)
                    if len(cells) >= 4:
                        safe_id_idx = id_idx if id_idx < len(cells) else 0
                        safe_name_idx = name_idx if name_idx < len(cells) else 1
                        safe_desc_idx = desc_idx if desc_idx < len(cells) else 2
                        safe_detail_idx = detail_idx if detail_idx < len(cells) else (3 if len(cells) > 3 else len(cells)-1)
                        
                        id_paras = extract_cell_styled_text(cells[safe_id_idx], ns)
                        name_paras = extract_cell_styled_text(cells[safe_name_idx], ns)
                        desc_paras = extract_cell_styled_text(cells[safe_desc_idx], ns)
                        detail_paras = extract_cell_styled_text(cells[safe_detail_idx], ns)
                        
                        col0 = get_plain_text(id_paras)
                        
                        is_header = "고유번호" in col0 or "요구사항" in col0 or "분류" in col0 or "No" in col0
                        has_valid_id = id_pattern.match(col0) or (len(col0) >= 3 and len(col0) <= 15 and not is_header)
                        
                        if has_valid_id and not is_header:
                            requirements.append({
                                'id': col0,
                                'id_styled': id_paras,
                                'name_styled': name_paras,
                                'desc_styled': desc_paras,
                                'detail_styled': detail_paras
                            })
        
        print(f"➜ 총 {len(requirements)}개의 요구사항 데이터를 성공적으로 추출했습니다.")
        return requirements

def fuse_to_hwpx(requirements):
    print("2. HWPX 양식 융합 작업 진행 중...")
    
    hp_ns_uri = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
    
    # 임시 작업용 폴더 생성
    temp_dir = r"C:\Users\KITC\Downloads\temp_hwpx_unzip"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    
    # 템플릿 복사 후 압축 해제
    with zipfile.ZipFile(hwpx_path, 'r') as z:
        z.extractall(temp_dir)
        
    header_xml_path = os.path.join(temp_dir, "Contents", "header.xml")
    section_xml_path = os.path.join(temp_dir, "Contents", "section0.xml")
    
    # 글자 모양 레지스트리 로드
    registry = HwpxCharPrRegistry(header_xml_path)
    
    # section0.xml 파싱
    ET.register_namespace('hp', hp_ns_uri)
    tree = ET.parse(section_xml_path)
    root = tree.getroot()
    
    # hp:tbl이 들어있는 최외곽 hp:p 찾기
    all_ps = root.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}p')
    template_p = None
    
    for p in all_ps:
        tbls = p.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}tbl')
        if len(tbls) > 0:
            template_p = p
            break
            
    if template_p is None:
        raise Exception("HWPX 양식에서 템플릿 요구사항 표 단락을 감지할 수 없습니다.")
        
    parent_map = {c: p for p in root.iter() for c in p}
    parent = parent_map[template_p]
    
    # 기존 template_p의 인덱스 찾고 부모에서 제거
    siblings = list(parent)
    template_idx = siblings.index(template_p)
    parent.remove(template_p)
    
    # 텍스트 및 서식 융합 주입 헬퍼 함수
    def fill_cell_with_style(tc, paragraphs):
        sub_list = tc.find('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}subList')
        if sub_list is None:
            return
        
        # 1. 기존 템플릿 스타일 참조 추출
        existing_ps = sub_list.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}p')
        first_p = existing_ps[0] if len(existing_ps) > 0 else None
        
        para_pr_id = first_p.get('paraPrIDRef') if first_p is not None else None
        style_id = first_p.get('styleIDRef') if first_p is not None else None
        
        base_char_pr_id = '0'
        if first_p is not None:
            first_run = first_p.find('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}run')
            if first_run is not None:
                base_char_pr_id = first_run.get('charPrIDRef', '0')
        
        # subList 내부의 자식 요소 전수 청소
        for child in list(sub_list):
            sub_list.remove(child)
            
        # 2. 런 단위로 한글 XML 노드 주입 가동
        for p_runs in paragraphs:
            new_p = ET.Element('{http://www.hancom.co.kr/hwpml/2011/paragraph}p')
            if para_pr_id:
                new_p.set('paraPrIDRef', para_pr_id)
            if style_id:
                new_p.set('styleIDRef', style_id)
                
            if not p_runs:
                # 빈 줄 처리
                run = ET.SubElement(new_p, '{http://www.hancom.co.kr/hwpml/2011/paragraph}run')
                run.set('charPrIDRef', base_char_pr_id)
                t = ET.SubElement(run, '{http://www.hancom.co.kr/hwpml/2011/paragraph}t')
                t.text = ""
            else:
                for run_data in p_runs:
                    if run_data['text'] == '\n':
                        # 문단 내 줄바꿈 (br)
                        run = ET.SubElement(new_p, '{http://www.hancom.co.kr/hwpml/2011/paragraph}run')
                        run.set('charPrIDRef', base_char_pr_id)
                        ET.SubElement(run, '{http://www.hancom.co.kr/hwpml/2011/paragraph}br')
                    else:
                        # 굵게/색상 동적 매핑된 charPr ID 가져오기
                        run_char_pr_id = registry.get_or_create_char_pr(base_char_pr_id, run_data['bold'], run_data['color'])
                        
                        run = ET.SubElement(new_p, '{http://www.hancom.co.kr/hwpml/2011/paragraph}run')
                        run.set('charPrIDRef', run_char_pr_id)
                        
                        t = ET.SubElement(run, '{http://www.hancom.co.kr/hwpml/2011/paragraph}t')
                        t.text = run_data['text']
                        
            sub_list.append(new_p)

    # 요구사항별 복제 및 융합 적용
    insert_idx = template_idx
    for idx, req in enumerate(requirements):
        cloned_p = ET.fromstring(ET.tostring(template_p))
        if 'id' in cloned_p.attrib:
            del cloned_p.attrib['id']
            
        if idx > 0:
            cloned_p.attrib['pageBreak'] = '1'
        else:
            cloned_p.attrib['pageBreak'] = '0'
            
        tbl = cloned_p.find('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}tbl')
        if tbl is not None:
            tc_list = tbl.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}tc')
            for tc in tc_list:
                t_nodes = tc.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}t')
                cell_text = "".join([t.text for t in t_nodes if t.text]).strip()
                
                if "{고유번호}" in cell_text or "{요구사항 고유번호}" in cell_text:
                    fill_cell_with_style(tc, req['id_styled'])
                elif "{요구사항 명칭}" in cell_text:
                    fill_cell_with_style(tc, req['name_styled'])
                elif "{정의}" in cell_text:
                    fill_cell_with_style(tc, req['desc_styled'])
                elif "{세부내용}" in cell_text:
                    fill_cell_with_style(tc, req['detail_styled'])
                    
        parent.insert(insert_idx, cloned_p)
        insert_idx += 1
        
        if idx < len(requirements) - 1:
            spacer_p = ET.Element('{http://www.hancom.co.kr/hwpml/2011/paragraph}p')
            spacer_p.attrib['pageBreak'] = '0'
            run = ET.SubElement(spacer_p, '{http://www.hancom.co.kr/hwpml/2011/paragraph}run')
            t = ET.SubElement(run, '{http://www.hancom.co.kr/hwpml/2011/paragraph}t')
            t.text = ""
            parent.insert(insert_idx, spacer_p)
            insert_idx += 1

    # XML 및 글자모양 갱신 저장
    tree.write(section_xml_path, encoding='utf-8', xml_declaration=True)
    registry.save()
    print("➜ section0.xml 및 header.xml 업데이트 완료.")

    # 패키징 압축
    print("3. HWPX 패키징 빌드 및 덮어쓰기 중...")
    if os.path.exists(output_path):
        os.remove(output_path)
        
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as target_zip:
        for root_dir, dirs, files in os.walk(temp_dir):
            for file in files:
                full_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(full_path, temp_dir)
                target_zip.write(full_path, rel_path)
                
    shutil.rmtree(temp_dir)
    print(f"➜ 최종 HWPX 양식 융합 파일이 안전하게 생성되었습니다:\n  [경로] {output_path}")

if __name__ == "__main__":
    try:
        reqs = parse_pptx_requirements()
        fuse_to_hwpx(reqs)
        print("\n🎉 [완료] PPT 요건기술서의 텍스트 세부 서식(Bold/색상)이 HWPX에 완벽하게 복제/이식되었습니다!")
    except Exception as e:
        print(f"\n❌ [에러 발생] 가공 중 중단되었습니다: {e}", file=sys.stderr)
