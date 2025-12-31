# HWP/HWPX Parser

한글 문서(.hwp, .hwpx) 파싱 및 레이아웃 정보 추출 라이브러리

## 주요 기능

- **텍스트 추출**: HWP/HWPX 문서에서 텍스트 추출
- **바운딩 박스 추출**: 문자, 문단, 표의 정확한 좌표 정보 (mm 단위)
- **표 구조화**: 제목/헤더/내용 분리로 LLM이 이해하기 쉬운 형태로 변환
- **계층적 구조**: 큰 제목 > 작은 제목 > 내용으로 문서 구조화
- **이미지 추출**: 임베디드 이미지 추출 및 좌표 정보 제공 (OCR 연동 지원)
- **WMF/EMF 변환**: 벡터 이미지를 PNG로 변환
- **시각화**: 바운딩 박스를 이미지/PDF에 그려서 확인 (단일/다중 페이지)
- **RAG 지원**: 청크 분할 및 메타데이터 포함

## 설치

```bash
pip install olefile Pillow
```

선택적 의존성 (이미지 추출 기능):
```bash
# WMF/EMF 변환을 위해 (선택)
brew install imagemagick  # macOS
apt-get install imagemagick  # Ubuntu/Debian
```

## 파일 구조

```
by_claude/
├── hwpx_parser.py          # HWPX 파싱 (XML 기반)
├── hwp_parser.py           # HWP 파싱 (OLE 기반)
├── document_extractor.py   # 구조화된 정보 추출 + 시각화
├── image_extractor.py      # 이미지 추출 (NEW!)
├── test_parsers.py         # 테스트 스크립트
└── results/                # 결과 저장 폴더
```

## 빠른 시작

```python
from hwpx_parser import parse_hwpx
from hwp_parser import parse_hwp
from document_extractor import extract_document_with_images, create_visualization_report

# 1. 문서 파싱
doc = parse_hwpx("document.hwpx")  # 또는 parse_hwp("document.hwp")

# 2. 구조화된 정보 + 이미지 추출
extracted = extract_document_with_images(
    doc,
    extract_images=True,           # 이미지 추출 활성화
    save_images_dir="output/images"  # 이미지 저장 디렉토리
)

# 3. 전체 리포트 생성 (이미지, JSON, 시각화 모두 포함)
create_visualization_report(extracted, "output/")
```

---

## 상세 사용법

### 1. 기본 파싱

#### HWPX 파싱

```python
from hwpx_parser import parse_hwpx

doc = parse_hwpx("document.hwpx")

# 기본 정보
print(f"제목: {doc.title}")
print(f"섹션 수: {len(doc.sections)}")

# 텍스트 추출
print(doc.to_text())        # 전체 텍스트
print(doc.to_markdown())    # 마크다운 변환
print(doc.to_json())        # JSON 변환

# 레이아웃 정보 포함 JSON
print(doc.to_json_with_layout())
```

#### HWP 파싱

```python
from hwp_parser import parse_hwp

doc = parse_hwp("document.hwp")

print(f"버전: {doc.header.version}")
print(f"압축: {doc.header.is_compressed}")
print(doc.to_text())
```

### 2. 레이아웃 정보 추출

```python
from hwpx_parser import parse_hwpx, extract_layout_elements

doc = parse_hwpx("document.hwpx")
elements, pages = extract_layout_elements(doc)

# 페이지 정보
for page in pages:
    print(f"페이지 {page.page_num + 1}: {page.width}mm x {page.height}mm")

# 요소 정보 (바운딩 박스 포함)
for elem in elements:
    print(f"[{elem.element_type}] ({elem.x:.1f}, {elem.y:.1f}) {elem.width:.1f}x{elem.height:.1f}mm")
    print(f"  텍스트: {elem.text[:50]}...")
```

### 3. 구조화된 정보 추출 (LLM/RAG용)

```python
from hwpx_parser import parse_hwpx
from document_extractor import extract_document

doc = parse_hwpx("document.hwpx")
extracted = extract_document(doc)

# 기본 정보
print(f"요소 수: {len(extracted.elements)}")
print(f"제목 수: {len(extracted.headings)}")
print(f"표 수: {len(extracted.tables)}")
print(f"계층 섹션 수: {len(extracted.hierarchical_sections)}")

# LLM에 적합한 구조화된 텍스트
print(extracted.to_structured_text())

# RAG용 청크 분할
chunks = extracted.to_rag_chunks(max_chunk_size=1000)
for chunk in chunks:
    print(f"--- 청크 ({len(chunk['text'])}자) ---")
    print(chunk["text"][:200])
    print(f"메타데이터: {chunk['metadata']}")

# JSON으로 저장
with open("extracted.json", "w", encoding="utf-8") as f:
    f.write(extracted.to_json())
```

---

## 🆕 이미지 추출

### 기본 이미지 추출

```python
from image_extractor import extract_images_from_hwp, extract_images_from_hwpx

# HWP 파일에서 이미지 추출
images = extract_images_from_hwp("document.hwp")

# HWPX 파일에서 이미지 추출
images = extract_images_from_hwpx("document.hwpx")

# 이미지 정보 확인
for img in images:
    print(f"이미지 ID: {img.bin_id}")
    print(f"파일명: {img.filename}")
    print(f"형식: {img.format}")
    print(f"크기: {len(img.data):,} bytes")
    print(f"해상도: {img.pixel_width}×{img.pixel_height} px")
    print(f"문서 내 위치: ({img.x:.1f}, {img.y:.1f}) mm")
    print(f"문서 내 크기: {img.width:.1f}×{img.height:.1f} mm")
    print(f"페이지: {img.page + 1}")
    print()
    
    # 이미지 저장
    img.save("output/images/")
```

### 이미지와 함께 문서 추출

```python
from hwp_parser import parse_hwp
from document_extractor import extract_document_with_images

doc = parse_hwp("document.hwp")

# 이미지 포함 추출
extracted = extract_document_with_images(
    doc,
    extract_images=True,              # 이미지 추출 활성화
    save_images_dir="output/images"   # 이미지 저장 디렉토리
)

# 추출된 이미지 확인
print(f"이미지 수: {len(extracted.images)}")
for img in extracted.images:
    print(f"  - {img.filename} ({img.format.upper()}, {img.file_size:,} bytes)")
    print(f"    위치: ({img.bbox.x:.1f}, {img.bbox.y:.1f}) mm")
    print(f"    페이지: {img.page + 1}")
```

### OCR 연동용 JSON 출력

이미지 추출 시 자동으로 생성되는 `*_images.json` 파일을 외부 OCR 서비스와 연동할 수 있습니다:

```json
{
  "document_title": "문서제목",
  "image_count": 2,
  "images": [
    {
      "image_id": "BIN0001",
      "filename": "BIN0001.jpg",
      "format": "jpg",
      "class": "image",
      "bbox_mm": {
        "x": 20.0,
        "y": 50.0,
        "width": 150.0,
        "height": 100.0
      },
      "bbox_px": {
        "width": 2481,
        "height": 3508
      },
      "page": 0,
      "saved_path": "/path/to/BIN0001.jpg",
      "ocr_text": "",
      "ocr_confidence": 0.0
    }
  ]
}
```

### OCR 연동 예시

```python
import json
from pathlib import Path

# 1. 이미지 메타데이터 로드
with open("output/document_images.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# 2. 각 이미지에 대해 OCR 수행
for img in data["images"]:
    image_path = img["saved_path"]
    
    # 외부 OCR 호출 (예: Tesseract, Cloud Vision API 등)
    # ocr_result = your_ocr_service.recognize(image_path)
    
    # 결과 저장
    # img["ocr_text"] = ocr_result["text"]
    # img["ocr_confidence"] = ocr_result["confidence"]

# 3. 업데이트된 메타데이터 저장
with open("output/document_images_with_ocr.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
```

### WMF/EMF 벡터 이미지 변환

```python
from image_extractor import extract_images_from_hwp

# WMF/EMF를 자동으로 PNG로 변환
images = extract_images_from_hwp("document.hwp")

for img in images:
    if img.format in ['wmf', 'emf']:
        print(f"벡터 이미지 발견: {img.filename}")
    
    # 변환과 함께 저장
    img.save("output/images/", convert_vector=True)
```

---

## 시각화

### 단일 페이지 시각화

```python
from document_extractor import visualize_elements

# 특정 페이지를 PNG로 저장
visualize_elements(
    extracted,
    "page_001.png",
    page_num=0,           # 페이지 번호 (0부터 시작)
    show_bbox=True,       # 바운딩 박스 표시
    show_text=True,       # 텍스트 표시
    scale=3.0,            # 확대 비율 (1mm = 3px)
    font_size=10,         # 폰트 크기
)
```

### 여러 페이지 시각화 (개별 PNG)

```python
from document_extractor import visualize_all_pages

# 모든 페이지를 개별 PNG 파일로 저장
saved_files = visualize_all_pages(
    extracted,
    "output_images/",     # 출력 디렉토리
    show_bbox=True,
    show_text=True,
    scale=3.0,
    font_size=10,
)

# 결과: output_images/문서제목_page_001.png, _page_002.png, ...
print(f"생성된 파일: {len(saved_files)}개")
for f in saved_files:
    print(f"  - {f}")
```

### 여러 페이지 시각화 (단일 PDF)

```python
from document_extractor import visualize_to_pdf

# 모든 페이지를 하나의 PDF로 저장
visualize_to_pdf(
    extracted,
    "output.pdf",         # 출력 PDF 경로
    show_bbox=True,
    show_text=True,
    scale=3.0,
    font_size=10,
)
```

### 전체 리포트 생성

```python
from document_extractor import create_visualization_report

# 이미지, JSON, 텍스트, 청크, 표 목록, 이미지 목록 모두 생성
saved_files = create_visualization_report(extracted, "output_dir/")

# 생성되는 파일들:
# - 문서제목_page_001.png, _page_002.png, ... (페이지별 이미지)
# - 문서제목_extracted.json (전체 추출 정보)
# - 문서제목_structured.txt (LLM용 구조화된 텍스트)
# - 문서제목_chunks.json (RAG용 청크)
# - 문서제목_tables.md (표 마크다운)
# - 문서제목_images.json (이미지 메타데이터, OCR 연동용)
# - 문서제목_images.md (이미지 목록 마크다운)
# - images/ (추출된 이미지 파일들)
```

---

## 출력 형식

### 레이아웃 JSON

```json
{
  "title": "문서제목",
  "unit_info": {
    "description": "좌표 단위는 mm"
  },
  "sections": [
    {
      "index": 0,
      "page": {
        "width_mm": 210.0,
        "height_mm": 297.0,
        "margins_mm": {"left": 20.0, "top": 20.0, "right": 20.0, "bottom": 20.0}
      },
      "paragraphs": [
        {
          "text": "문단 텍스트",
          "bbox": {"x": 20.0, "y": 30.0, "width": 170.0, "height": 5.0, "x2": 190.0, "y2": 35.0},
          "line_segments": [
            {"x_mm": 20.0, "y_mm": 30.0, "width_mm": 170.0, "height_mm": 5.0}
          ]
        }
      ]
    }
  ]
}
```

### 구조화된 텍스트 (LLM용)

```
# 문서 제목

[문서 유형] HWPX
[페이지 수] 3

## 문서 내용

### 가. 첫 번째 섹션

내용 텍스트...

[표 제목] 표 1
[헤더] 열1 | 열2 | 열3
[행 1] 데이터1 | 데이터2 | 데이터3
```

### RAG 청크

```json
[
  {
    "text": "## 섹션 제목\n\n내용...",
    "metadata": {
      "title": "섹션 제목",
      "level": 2,
      "page": 0,
      "source": "document.hwpx"
    }
  }
]
```

---

## 완전한 예시: 문서 파싱부터 OCR 연동까지

```python
from pathlib import Path
from hwp_parser import parse_hwp
from document_extractor import extract_document_with_images, create_visualization_report

# 1. 문서 파싱
hwp_file = Path("document.hwp")
doc = parse_hwp(hwp_file)

# 2. 이미지 포함 구조화된 정보 추출
output_dir = Path("output")
extracted = extract_document_with_images(
    doc,
    extract_images=True,
    save_images_dir=output_dir / "images"
)

# 3. 전체 리포트 생성
saved_files = create_visualization_report(extracted, output_dir)

print(f"✅ 총 {len(saved_files)}개 파일 생성")
print(f"  - 텍스트 요소: {len(extracted.elements)}개")
print(f"  - 표: {len(extracted.tables)}개")
print(f"  - 이미지: {len(extracted.images)}개")
print(f"  - 페이지: {len(extracted.pages)}개")

# 4. 이미지 정보 확인
for img in extracted.images:
    print(f"\n이미지: {img.filename}")
    print(f"  - 형식: {img.format.upper()}")
    print(f"  - 크기: {img.file_size:,} bytes")
    print(f"  - 해상도: {img.pixel_width}×{img.pixel_height} px")
    print(f"  - 위치: ({img.bbox.x:.1f}, {img.bbox.y:.1f}) mm")
    print(f"  - 페이지: {img.page + 1}")
    print(f"  - 저장 경로: {img.saved_path}")

# 5. (선택) 외부 OCR 연동
import json

images_json_path = output_dir / f"{extracted.title}_images.json"
with open(images_json_path, "r", encoding="utf-8") as f:
    images_data = json.load(f)

for img_data in images_data["images"]:
    image_path = img_data["saved_path"]
    
    # 여기에 OCR 로직 추가
    # ocr_result = your_ocr_api(image_path)
    # img_data["ocr_text"] = ocr_result["text"]
    # img_data["ocr_confidence"] = ocr_result["confidence"]

# 업데이트된 JSON 저장
with open(images_json_path, "w", encoding="utf-8") as f:
    json.dump(images_data, f, ensure_ascii=False, indent=2)

print(f"\n✅ OCR 연동 준비 완료: {images_json_path}")
```

---

## 좌표 단위

| 단위 | 설명 |
|------|------|
| HWPUNIT | HWP 내부 단위 (1 HWPUNIT = 1/7200 인치) |
| mm | 출력 단위 (1 HWPUNIT ≈ 0.00353mm) |

좌표 원점은 **페이지 왼쪽 상단**입니다.

## 시각화 색상 범례

| 요소 유형 | 색상 |
|----------|------|
| heading (제목) | 분홍색 (#E91E63) |
| paragraph (문단) | 파란색 (#2196F3) |
| table (표) | 녹색 (#4CAF50) |
| table_cell (표 셀) | 주황색 (#FF9800) |

## 제목 인식 패턴

다음 패턴의 텍스트를 제목으로 인식합니다:

| 레벨 | 패턴 예시 |
|------|----------|
| 1 (대제목) | 제1장, 제1편, Ⅰ., Ⅱ., Ⅲ. |
| 2 (중제목) | 가., 1., 【제목】, [제목], 1) |
| 3 (소제목) | 가), ①, ②, ③ |

---

## 테스트 실행

```bash
cd by_claude
python document_extractor.py  # 전체 파이프라인 테스트
python image_extractor.py     # 이미지 추출 테스트
```

## 의존성

| 패키지 | 용도 | 필수 |
|--------|------|------|
| olefile | HWP 파일 파싱 | HWP 파싱시 필수 |
| Pillow | 시각화 이미지/PDF 생성, 이미지 추출 | 시각화 또는 이미지 추출 시 필수 |
| ImageMagick | WMF/EMF → PNG 변환 | WMF/EMF 변환 시 선택 |
| LibreOffice | WMF/EMF → PNG 변환 (대안) | WMF/EMF 변환 시 선택 |

## 지원 형식

### 문서 형식

| 형식 | 확장자 | 설명 |
|------|--------|------|
| HWPX | .hwpx | 한글 2014 이후 XML 기반 형식 (ZIP 압축) |
| HWP | .hwp | 한글 97 이후 OLE 기반 형식 |

### 이미지 형식

| 형식 | 지원 | 비고 |
|------|------|------|
| JPEG | ✅ | 직접 추출 |
| PNG | ✅ | 직접 추출 |
| BMP | ✅ | zlib 압축 해제 지원 |
| GIF | ✅ | 직접 추출 |
| WMF | ✅ | PNG 변환 필요 (ImageMagick/LibreOffice) |
| EMF | ✅ | PNG 변환 필요 (ImageMagick/LibreOffice) |

## 제한사항

- 암호화된 문서는 지원하지 않습니다
- 복잡한 레이아웃 (다단, 텍스트 상자 등)의 좌표가 부정확할 수 있습니다
- HWP의 경우 표 파싱이 제한적입니다
- HWP 이미지 좌표는 GSO 컨트롤 파싱에 의존하며, 일부 문서에서 부정확할 수 있습니다
- WMF/EMF 변환은 외부 도구(ImageMagick/LibreOffice)가 설치되어 있어야 합니다

## 문제 해결

### 이미지가 추출되지 않는 경우
- `olefile`과 `Pillow`가 설치되어 있는지 확인하세요
- HWP 파일이 암호화되어 있지 않은지 확인하세요

### WMF/EMF 이미지가 변환되지 않는 경우
```bash
# ImageMagick 설치 확인
which convert

# LibreOffice 설치 확인 (macOS)
ls /Applications/LibreOffice.app/Contents/MacOS/soffice
```

### 이미지 좌표가 0으로 나오는 경우
- HWP의 GSO 파싱은 복잡하며, 일부 문서에서는 좌표 추출이 제한적입니다
- 이미지 파일 자체는 정상적으로 추출되며, 픽셀 크기 정보는 제공됩니다

## 라이선스

MIT License
