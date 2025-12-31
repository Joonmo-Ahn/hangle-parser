"""
HWPX/HWP 파서 테스트 스크립트

이 스크립트는 다음을 테스트합니다:
1. HWPX 파일 파싱 및 레이아웃 추출
2. HWP 파일 파싱 및 레이아웃 추출
3. 문서 구조화 (제목, 표, 문맥)
4. 시각화 출력
"""

import sys
from pathlib import Path

# 현재 디렉토리를 경로에 추가
sys.path.insert(0, str(Path(__file__).parent))

from hwpx_parser import parse_hwpx, extract_layout_elements
from hwp_parser import parse_hwp, extract_layout_elements as extract_hwp_layout
from document_extractor import extract_document, create_visualization_report


def test_hwpx(file_path: Path, output_dir: Path):
    """HWPX 파일 테스트"""
    print(f"\n{'='*70}")
    print(f"HWPX 파일 테스트: {file_path.name}")
    print('='*70)

    # 파싱
    print("\n1. 파싱 중...")
    doc = parse_hwpx(file_path)

    print(f"   - 제목: {doc.title}")
    print(f"   - 버전: {doc.version.application} {doc.version.app_version}")
    print(f"   - 섹션 수: {len(doc.sections)}")

    for section in doc.sections:
        page_mm = section.page_props.to_mm()
        print(f"   - Section {section.index + 1}: {page_mm['width_mm']}mm × {page_mm['height_mm']}mm")
        print(f"     문단 수: {len(section.paragraphs)}")

    # 텍스트 추출
    print("\n2. 텍스트 추출:")
    text = doc.to_text()
    print(f"   - 전체 텍스트 길이: {len(text)}자")
    print(f"   - 첫 200자: {text[:200]}...")

    # 레이아웃 추출
    print("\n3. 레이아웃 추출:")
    elements, pages = extract_layout_elements(doc)
    print(f"   - 페이지 수: {len(pages)}")
    print(f"   - 레이아웃 요소 수: {len(elements)}")

    # 좌표 샘플
    print("\n4. 좌표 샘플 (처음 10개):")
    for elem in elements[:10]:
        print(f"   [{elem.element_type:10}] ({elem.x:6.1f}, {elem.y:6.1f}) {elem.width:6.1f}×{elem.height:5.1f}mm")
        if elem.text:
            print(f"              텍스트: {elem.text[:40]}...")

    # 구조화된 정보 추출
    print("\n5. 구조화된 정보 추출:")
    extracted = extract_document(doc)
    print(f"   - 요소 수: {len(extracted.elements)}")
    print(f"   - 제목 수: {len(extracted.headings)}")
    print(f"   - 문단 수: {len(extracted.paragraphs)}")
    print(f"   - 표 수: {len(extracted.tables)}")
    print(f"   - 계층 섹션 수: {len(extracted.hierarchical_sections)}")

    # 제목 목록
    if extracted.headings:
        print("\n   제목 목록:")
        for h in extracted.headings[:5]:
            print(f"     [L{h.level}] {h.text[:50]}...")

    # 표 정보
    if extracted.tables:
        print("\n   표 목록:")
        for t in extracted.tables:
            print(f"     - {t.title or '(제목 없음)'} ({t.row_count}×{t.col_count})")

    # 결과 저장
    print("\n6. 결과 저장:")
    output_subdir = output_dir / "hwpx"
    output_subdir.mkdir(parents=True, exist_ok=True)

    saved_files = create_visualization_report(extracted, output_subdir)
    for f in saved_files:
        print(f"   - {f}")

    # JSON with layout
    json_layout_path = output_subdir / f"{doc.title}_layout.json"
    with open(json_layout_path, "w", encoding="utf-8") as f:
        f.write(doc.to_json_with_layout())
    print(f"   - {json_layout_path}")

    return extracted


def test_hwp(file_path: Path, output_dir: Path):
    """HWP 파일 테스트"""
    print(f"\n{'='*70}")
    print(f"HWP 파일 테스트: {file_path.name}")
    print('='*70)

    # 파싱
    print("\n1. 파싱 중...")
    doc = parse_hwp(file_path)

    print(f"   - 제목: {doc.title}")
    print(f"   - 버전: {doc.header.version}")
    print(f"   - 압축: {'예' if doc.header.is_compressed else '아니오'}")
    print(f"   - 섹션 수: {len(doc.sections)}")
    print(f"   - 글꼴 수: {len(doc.fonts)}")

    for section in doc.sections:
        print(f"   - Section {section.index + 1}: {section.page_width_mm():.1f}mm × {section.page_height_mm():.1f}mm")
        print(f"     문단 수: {len(section.paragraphs)}")

    # 텍스트 추출
    print("\n2. 텍스트 추출:")
    text = doc.to_text()
    print(f"   - 전체 텍스트 길이: {len(text)}자")
    print(f"   - 첫 200자: {text[:200]}...")

    # 레이아웃 추출
    print("\n3. 레이아웃 추출:")
    elements, pages = extract_hwp_layout(doc)
    print(f"   - 페이지 수: {len(pages)}")
    print(f"   - 레이아웃 요소 수: {len(elements)}")

    # 좌표 샘플
    print("\n4. 좌표 샘플 (처음 10개):")
    for elem in elements[:10]:
        print(f"   [{elem.element_type:10}] ({elem.x:6.1f}, {elem.y:6.1f}) {elem.width:6.1f}×{elem.height:5.1f}mm")
        if elem.text:
            print(f"              텍스트: {elem.text[:40]}...")

    # 구조화된 정보 추출
    print("\n5. 구조화된 정보 추출:")
    extracted = extract_document(doc)
    print(f"   - 요소 수: {len(extracted.elements)}")
    print(f"   - 제목 수: {len(extracted.headings)}")
    print(f"   - 문단 수: {len(extracted.paragraphs)}")
    print(f"   - 표 수: {len(extracted.tables)}")
    print(f"   - 계층 섹션 수: {len(extracted.hierarchical_sections)}")

    # 제목 목록
    if extracted.headings:
        print("\n   제목 목록:")
        for h in extracted.headings[:5]:
            print(f"     [L{h.level}] {h.text[:50]}...")

    # 결과 저장
    print("\n6. 결과 저장:")
    output_subdir = output_dir / "hwp"
    output_subdir.mkdir(parents=True, exist_ok=True)

    saved_files = create_visualization_report(extracted, output_subdir)
    for f in saved_files:
        print(f"   - {f}")

    # JSON with layout
    json_layout_path = output_subdir / f"{doc.title}_layout.json"
    with open(json_layout_path, "w", encoding="utf-8") as f:
        f.write(doc.to_json_with_layout())
    print(f"   - {json_layout_path}")

    return extracted


def main():
    """메인 함수"""
    print("="*70)
    print("HWPX/HWP 파서 테스트")
    print("="*70)

    # 경로 설정
    data_dir = Path(__file__).parent.parent / "data" / "docs"
    output_dir = Path(__file__).parent / "results"
    output_dir.mkdir(parents=True, exist_ok=True)

    hwpx_file = data_dir / "은행권 광고심의 결과 보고서(양식)vF (1).hwpx"
    hwp_file = data_dir / "2. [농협] 광고안(B).hwp"

    results = {}

    # HWPX 테스트
    if hwpx_file.exists():
        try:
            results['hwpx'] = test_hwpx(hwpx_file, output_dir)
        except Exception as e:
            print(f"\nHWPX 테스트 오류: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"\nHWPX 파일을 찾을 수 없습니다: {hwpx_file}")

    # HWP 테스트
    if hwp_file.exists():
        try:
            results['hwp'] = test_hwp(hwp_file, output_dir)
        except Exception as e:
            print(f"\nHWP 테스트 오류: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"\nHWP 파일을 찾을 수 없습니다: {hwp_file}")

    # 요약
    print(f"\n{'='*70}")
    print("테스트 요약")
    print('='*70)

    if 'hwpx' in results:
        ex = results['hwpx']
        print(f"\nHWPX 결과:")
        print(f"  - 총 요소: {len(ex.elements)}")
        print(f"  - 유효한 좌표를 가진 요소: {len([e for e in ex.elements if e.bbox.is_valid()])}")
        print(f"  - 제목: {len(ex.headings)}")
        print(f"  - 표: {len(ex.tables)}")

    if 'hwp' in results:
        ex = results['hwp']
        print(f"\nHWP 결과:")
        print(f"  - 총 요소: {len(ex.elements)}")
        print(f"  - 유효한 좌표를 가진 요소: {len([e for e in ex.elements if e.bbox.is_valid()])}")
        print(f"  - 제목: {len(ex.headings)}")
        print(f"  - 표: {len(ex.tables)}")

    print(f"\n결과 저장 위치: {output_dir}")
    print("\n테스트 완료!")


if __name__ == "__main__":
    main()
