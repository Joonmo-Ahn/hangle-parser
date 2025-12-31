"""
📄 문서 추출 및 시각화 데모

이 스크립트는 HWPX/HWP 파일에서 구조화된 정보를 추출하고 시각화하는 방법을 보여줍니다.

사용 순서:
1. 문서 파싱 (parse_hwpx 또는 parse_hwp)
2. 요소 추출 (extract_document_elements)
3. 시각화 (visualize_elements)

LLM/RAG용 활용:
- extracted.to_structured_text() : 문맥에 맞는 구조화된 텍스트
- extracted.tables : 표 목록 (제목/헤더/내용 분리)
- extracted.headings : 제목 목록
- extracted.get_full_text() : 전체 텍스트
"""

from __future__ import annotations
from pathlib import Path
from typing import Union
import json
import sys

# 경로 설정
sys.path.insert(0, str(Path(__file__).parent))

from hwpx_parser_cursor import parse_hwpx
from hwp_parser_cursor import parse_hwp
from document_extractor import (
    extract_document_elements, 
    visualize_elements,
    create_visualization_report,
    ExtractedDocument,
    TableStructure,
)

def demo_hwpx(file_path: Union[str, Path], output_dir: Union[str, Path]):
    """HWPX 문서 처리 데모"""
    print("=" * 70)
    print(f"📄 HWPX 문서 처리: {Path(file_path).name}")
    print("=" * 70)
    
    # Step 1: 파싱
    print("\n1️⃣ 문서 파싱 중...")
    doc = parse_hwpx(file_path)
    print(f"   ✅ 파싱 완료: {len(doc.sections)} 섹션")
    
    # Step 2: 요소 추출
    print("\n2️⃣ 요소 추출 중...")
    extracted = extract_document_elements(doc)
    print(f"   ✅ 추출 완료:")
    print(f"      - 총 요소: {len(extracted.elements)}개")
    print(f"      - 제목: {len(extracted.headings)}개")
    print(f"      - 문단: {len(extracted.paragraphs)}개")
    print(f"      - 표: {len(extracted.tables)}개")
    
    # Step 3: 좌표 정보 샘플 출력
    print("\n3️⃣ 좌표 정보 샘플 (처음 5개):")
    for elem in extracted.elements[:5]:
        print(f"   [{elem.element_type}] {elem.text[:30]}...")
        print(f"      위치: ({elem.bbox.x:.1f}, {elem.bbox.y:.1f}) mm")
        print(f"      크기: {elem.bbox.width:.1f} × {elem.bbox.height:.1f} mm")
    
    # Step 4: 표 구조 출력
    if extracted.tables:
        print("\n4️⃣ 표 구조 (처음 2개):")
        for i, table in enumerate(extracted.tables[:2]):
            print(f"\n   📊 표 {i+1}: {table.title[:30]}...")
            print(f"      크기: {table.bbox.width:.1f} × {table.bbox.height:.1f} mm")
            print(f"      헤더: {len(table.headers)} 행")
            print(f"      데이터: {len(table.rows)} 행")
            
            # LLM용 구조화된 텍스트
            structured = table.to_structured_text()
            print(f"      LLM용 텍스트 미리보기:")
            for line in structured.split('\n')[:3]:
                print(f"         {line}")
    
    # Step 5: 시각화
    print("\n5️⃣ 시각화 생성 중...")
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 단일 페이지 시각화
    img_path = output_dir / f"{Path(file_path).stem}_visualized.png"
    visualize_elements(extracted, img_path, page_num=0)
    
    # Step 6: 결과 저장
    print("\n6️⃣ 결과 저장 중...")
    
    # LLM용 구조화된 텍스트
    txt_path = output_dir / f"{Path(file_path).stem}_for_llm.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(extracted.to_structured_text())
    print(f"   ✅ LLM용 텍스트: {txt_path}")
    
    # JSON (전체 정보)
    json_path = output_dir / f"{Path(file_path).stem}_elements.json"
    with open(json_path, "w", encoding="utf-8") as f:
        f.write(extracted.to_json(indent=2))
    print(f"   ✅ JSON: {json_path}")
    
    return extracted


def demo_hwp(file_path: Union[str, Path], output_dir: Union[str, Path]):
    """HWP 문서 처리 데모"""
    print("=" * 70)
    print(f"📄 HWP 문서 처리: {Path(file_path).name}")
    print("=" * 70)
    
    # Step 1: 파싱
    print("\n1️⃣ 문서 파싱 중...")
    doc = parse_hwp(file_path)
    print(f"   ✅ 파싱 완료: {len(doc.sections)} 섹션")
    
    # Step 2: 요소 추출
    print("\n2️⃣ 요소 추출 중...")
    extracted = extract_document_elements(doc)
    print(f"   ✅ 추출 완료:")
    print(f"      - 총 요소: {len(extracted.elements)}개")
    print(f"      - 제목: {len(extracted.headings)}개")
    print(f"      - 문단: {len(extracted.paragraphs)}개")
    print(f"      - 표: {len(extracted.tables)}개")
    
    # Step 3: 좌표 정보 샘플 출력
    print("\n3️⃣ 좌표 정보 샘플 (처음 5개):")
    for elem in extracted.elements[:5]:
        print(f"   [{elem.element_type}] {elem.text[:30]}...")
        print(f"      위치: ({elem.bbox.x:.1f}, {elem.bbox.y:.1f}) mm")
        print(f"      크기: {elem.bbox.width:.1f} × {elem.bbox.height:.1f} mm")
    
    # Step 4: 시각화
    print("\n4️⃣ 시각화 생성 중...")
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    img_path = output_dir / f"{Path(file_path).stem}_visualized.png"
    visualize_elements(extracted, img_path, page_num=0)
    
    # Step 5: 결과 저장
    print("\n5️⃣ 결과 저장 중...")
    
    # LLM용 구조화된 텍스트
    txt_path = output_dir / f"{Path(file_path).stem}_for_llm.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(extracted.to_structured_text())
    print(f"   ✅ LLM용 텍스트: {txt_path}")
    
    # JSON
    json_path = output_dir / f"{Path(file_path).stem}_elements.json"
    with open(json_path, "w", encoding="utf-8") as f:
        f.write(extracted.to_json(indent=2))
    print(f"   ✅ JSON: {json_path}")
    
    return extracted


def print_usage_examples():
    """사용 예시 출력"""
    print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║                      📖 사용 예시 (코드에서 직접 사용)                        ║
╚══════════════════════════════════════════════════════════════════════════════╝

# 1. 기본 사용법

from hwpx_parser_cursor import parse_hwpx
from hwp_parser_cursor import parse_hwp
from document_extractor import extract_document_elements, visualize_elements

# HWPX 파싱 및 추출
doc = parse_hwpx("document.hwpx")
extracted = extract_document_elements(doc)

# HWP 파싱 및 추출
doc = parse_hwp("document.hwp")
extracted = extract_document_elements(doc)


# 2. 추출된 정보 활용

# 전체 텍스트
full_text = extracted.get_full_text()

# LLM용 구조화된 텍스트
structured_text = extracted.to_structured_text()

# 제목 목록
for heading in extracted.headings:
    print(f"{'  ' * heading.level}{heading.text}")

# 표 목록 (제목/헤더/내용 분리)
for table in extracted.tables:
    print(f"표 제목: {table.title}")
    print(f"헤더: {table.headers}")
    print(f"데이터: {table.rows}")
    print(table.to_markdown())
    print(table.to_structured_text())


# 3. 개별 요소의 바운딩 박스

for elem in extracted.elements:
    print(f"요소: {elem.element_type}")
    print(f"텍스트: {elem.text}")
    print(f"위치: x={elem.bbox.x}mm, y={elem.bbox.y}mm")
    print(f"크기: {elem.bbox.width}mm × {elem.bbox.height}mm")
    print(f"페이지: {elem.page + 1}")


# 4. 시각화

# 단일 페이지
visualize_elements(extracted, "output.png", page_num=0)

# 전체 문서 리포트
from document_extractor import create_visualization_report
create_visualization_report(extracted, "output_dir/")


# 5. JSON 저장 및 로드

import json

# 저장
with open("extracted.json", "w") as f:
    f.write(extracted.to_json())

# 로드 (dict로)
with open("extracted.json", "r") as f:
    data = json.load(f)

""")


if __name__ == "__main__":
    data_dir = Path(__file__).parent.parent / "data"
    output_dir = Path(__file__).parent / "results" / "demo"
    
    hwpx_file = data_dir / "은행권 광고심의 결과 보고서(양식)vF (1).hwpx"
    hwp_file = data_dir / "2. [농협] 광고안(B).hwp"
    
    # 사용 예시 출력
    print_usage_examples()
    
    # HWPX 데모
    if hwpx_file.exists():
        demo_hwpx(hwpx_file, output_dir)
    
    print("\n" + "=" * 70)
    
    # HWP 데모
    if hwp_file.exists():
        demo_hwp(hwp_file, output_dir)
    
    print("\n✅ 데모 완료!")
    print(f"   결과 폴더: {output_dir}")

