"""
HWPX 및 HWP 파서 검증 스크립트

이 스크립트는 두 파서가 올바르게 동작하는지 검증합니다.

테스트 파일:
1. HWPX: /Users/jmahn/Project/code/hwp/data/은행권 광고심의 결과 보고서(양식)vF (1).hwpx
2. HWP: /Users/jmahn/Project/code/hwp/data/2. [농협] 광고안(B).hwp

결과 저장 위치: /Users/jmahn/Project/code/hwp/by_cursor/results/
"""

import sys
import json
from pathlib import Path

# 현재 디렉토리를 모듈 경로에 추가
sys.path.insert(0, str(Path(__file__).parent))

from hwpx_parser_cursor import (
    parse_hwpx, 
    extract_layout_elements, 
    extract_layout_summary,
    visualize_document_pil,
)
from hwp_parser_cursor import parse_hwp


def test_hwpx_parser(hwpx_file: str, output_dir: Path):
    """HWPX 파서 테스트"""
    print("\n" + "=" * 70)
    print("📄 HWPX 파서 테스트")
    print("=" * 70)
    print(f"파일: {hwpx_file}")
    
    try:
        # 1. 파싱
        print("\n1️⃣ 파싱 중...")
        doc = parse_hwpx(hwpx_file)
        print(f"   ✅ 파싱 성공!")
        
        # 2. 기본 정보 출력
        print(f"\n2️⃣ 문서 정보:")
        print(f"   - 제목: {doc.title}")
        print(f"   - 버전: {doc.version.application} {doc.version.app_version}")
        print(f"   - 섹션 수: {len(doc.sections)}")
        
        total_paras = sum(len(s.paragraphs) for s in doc.sections)
        total_tables = sum(sum(len(p.tables) for p in s.paragraphs) for s in doc.sections)
        print(f"   - 총 문단 수: {total_paras}")
        print(f"   - 총 테이블 수: {total_tables}")
        
        for section in doc.sections:
            page_mm = section.page_props.to_mm()
            print(f"   - Section {section.index + 1}: {page_mm['width_mm']}mm × {page_mm['height_mm']}mm ({page_mm['orientation']})")
        
        # 3. 레이아웃 요소 추출
        print(f"\n3️⃣ 레이아웃 요소 추출...")
        elements, pages = extract_layout_elements(doc)
        print(f"   - 페이지 수: {len(pages)}")
        print(f"   - 요소 수: {len(elements)}")
        
        text_count = sum(1 for e in elements if e.element_type == "text")
        table_count = sum(1 for e in elements if e.element_type == "table")
        cell_count = sum(1 for e in elements if e.element_type == "table_cell")
        print(f"   - 텍스트 요소: {text_count}")
        print(f"   - 테이블 요소: {table_count}")
        print(f"   - 테이블 셀: {cell_count}")
        
        # 4. 결과 저장
        print(f"\n4️⃣ 결과 저장 중...")
        
        # 텍스트 저장
        txt_file = output_dir / f"{doc.title}_extracted.txt"
        with open(txt_file, "w", encoding="utf-8") as f:
            f.write(doc.to_text())
        print(f"   ✅ 텍스트: {txt_file.name}")
        
        # 마크다운 저장
        md_file = output_dir / f"{doc.title}_parsed.md"
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(doc.to_markdown())
        print(f"   ✅ 마크다운: {md_file.name}")
        
        # 레이아웃 마크다운 저장
        md_layout_file = output_dir / f"{doc.title}_layout.md"
        with open(md_layout_file, "w", encoding="utf-8") as f:
            f.write(doc.to_markdown_with_layout())
        print(f"   ✅ 레이아웃 마크다운: {md_layout_file.name}")
        
        # JSON 저장
        json_file = output_dir / f"{doc.title}_parsed.json"
        with open(json_file, "w", encoding="utf-8") as f:
            f.write(doc.to_json())
        print(f"   ✅ JSON: {json_file.name}")
        
        # 레이아웃 JSON 저장
        json_layout_file = output_dir / f"{doc.title}_layout.json"
        with open(json_layout_file, "w", encoding="utf-8") as f:
            f.write(doc.to_json_with_layout())
        print(f"   ✅ 레이아웃 JSON: {json_layout_file.name}")
        
        # 레이아웃 요소 JSON 저장
        summary = extract_layout_summary(doc)
        elements_file = output_dir / f"{doc.title}_elements.json"
        with open(elements_file, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        print(f"   ✅ 레이아웃 요소: {elements_file.name}")
        
        # 시각화 이미지 저장
        try:
            img_file = output_dir / f"{doc.title}_visualization.png"
            visualize_document_pil(doc, img_file, scale=3.0)
            print(f"   ✅ 시각화: {img_file.name}")
        except Exception as e:
            print(f"   ⚠️ 시각화 실패: {e}")
        
        # 5. 텍스트 미리보기
        print(f"\n5️⃣ 텍스트 미리보기 (처음 500자):")
        print("-" * 50)
        text = doc.to_text()
        print(text[:500] if len(text) > 500 else text)
        print("-" * 50)
        
        print(f"\n✅ HWPX 파서 테스트 완료!")
        return True
        
    except Exception as e:
        print(f"\n❌ HWPX 파서 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_hwp_parser(hwp_file: str, output_dir: Path):
    """HWP 파서 테스트"""
    print("\n" + "=" * 70)
    print("📄 HWP 파서 테스트")
    print("=" * 70)
    print(f"파일: {hwp_file}")
    
    try:
        # 1. 파싱
        print("\n1️⃣ 파싱 중...")
        doc = parse_hwp(hwp_file)
        print(f"   ✅ 파싱 성공!")
        
        # 2. 기본 정보 출력
        print(f"\n2️⃣ 문서 정보:")
        print(f"   - 제목: {doc.title}")
        print(f"   - 버전: {doc.header.version}")
        print(f"   - 압축: {'예' if doc.header.is_compressed else '아니오'}")
        print(f"   - 암호화: {'예' if doc.header.is_encrypted else '아니오'}")
        print(f"   - 섹션 수: {len(doc.sections)}")
        print(f"   - 글꼴 수: {len(doc.fonts)}")
        
        total_paras = sum(len(s.paragraphs) for s in doc.sections)
        print(f"   - 총 문단 수: {total_paras}")
        
        if doc.fonts:
            print(f"   - 글꼴 목록:")
            for font in doc.fonts[:5]:
                print(f"     · {font.name}")
            if len(doc.fonts) > 5:
                print(f"     · ... 외 {len(doc.fonts) - 5}개")
        
        # 3. 결과 저장
        print(f"\n3️⃣ 결과 저장 중...")
        
        # 텍스트 저장
        txt_file = output_dir / f"{doc.title}_extracted.txt"
        with open(txt_file, "w", encoding="utf-8") as f:
            f.write(doc.to_text())
        print(f"   ✅ 텍스트: {txt_file.name}")
        
        # 마크다운 저장
        md_file = output_dir / f"{doc.title}_parsed.md"
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(doc.to_markdown())
        print(f"   ✅ 마크다운: {md_file.name}")
        
        # JSON 저장
        json_file = output_dir / f"{doc.title}_parsed.json"
        with open(json_file, "w", encoding="utf-8") as f:
            f.write(doc.to_json())
        print(f"   ✅ JSON: {json_file.name}")
        
        # 4. 텍스트 미리보기
        print(f"\n4️⃣ 텍스트 미리보기 (처음 500자):")
        print("-" * 50)
        text = doc.to_text()
        print(text[:500] if len(text) > 500 else text)
        print("-" * 50)
        
        print(f"\n✅ HWP 파서 테스트 완료!")
        return True
        
    except Exception as e:
        print(f"\n❌ HWP 파서 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """메인 함수"""
    print("=" * 70)
    print("🔍 HWPX & HWP 파서 검증 스크립트")
    print("=" * 70)
    
    # 경로 설정
    base_dir = Path(__file__).parent.parent
    data_dir = base_dir / "data" / "docs"
    output_dir = Path(__file__).parent / "results"
    
    # 테스트 파일
    hwpx_file = data_dir / "은행권 광고심의 결과 보고서(양식)vF (1).hwpx"
    hwp_file = data_dir / "2. [농협] 광고안(B).hwp"
    
    # 출력 디렉토리 생성
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"\n📁 데이터 디렉토리: {data_dir}")
    print(f"📁 출력 디렉토리: {output_dir}")
    
    # 파일 존재 확인
    print(f"\n📋 테스트 파일 확인:")
    print(f"   - HWPX: {hwpx_file.name} {'✅ 존재' if hwpx_file.exists() else '❌ 없음'}")
    print(f"   - HWP: {hwp_file.name} {'✅ 존재' if hwp_file.exists() else '❌ 없음'}")
    
    results = []
    
    # HWPX 테스트
    if hwpx_file.exists():
        results.append(("HWPX", test_hwpx_parser(str(hwpx_file), output_dir)))
    else:
        print(f"\n⚠️ HWPX 파일이 없어 테스트를 건너뜁니다.")
        results.append(("HWPX", None))
    
    # HWP 테스트
    if hwp_file.exists():
        results.append(("HWP", test_hwp_parser(str(hwp_file), output_dir)))
    else:
        print(f"\n⚠️ HWP 파일이 없어 테스트를 건너뜁니다.")
        results.append(("HWP", None))
    
    # 결과 요약
    print("\n" + "=" * 70)
    print("📊 테스트 결과 요약")
    print("=" * 70)
    
    for name, result in results:
        if result is True:
            status = "✅ 성공"
        elif result is False:
            status = "❌ 실패"
        else:
            status = "⏭️ 건너뜀"
        print(f"   {name}: {status}")
    
    print(f"\n📁 결과 파일 위치: {output_dir}")
    print("=" * 70)
    
    # 생성된 파일 목록
    print("\n📋 생성된 파일:")
    for f in sorted(output_dir.iterdir()):
        size = f.stat().st_size
        if size > 1024 * 1024:
            size_str = f"{size / 1024 / 1024:.1f} MB"
        elif size > 1024:
            size_str = f"{size / 1024:.1f} KB"
        else:
            size_str = f"{size} B"
        print(f"   - {f.name} ({size_str})")


if __name__ == "__main__":
    main()

