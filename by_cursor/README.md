# 📄 HWP/HWPX Parser for Python

한글 문서 파일(`.hwp`, `.hwpx`)을 파싱하여 텍스트, 레이아웃, 좌표 정보, 이미지를 추출하는 Python 라이브러리입니다.

**LLM/RAG 시스템에 최적화된 구조화된 출력**을 제공합니다.

---

## ✨ 주요 기능

| 기능 | 설명 |
|------|------|
| 📝 **텍스트 추출** | 문단, 표, 제목 등 모든 텍스트 추출 |
| 📐 **레이아웃 정보** | 바운딩 박스, 좌표, 크기 (mm 단위) |
| 📊 **표 구조화** | 제목/헤더/내용 분리 (LLM 친화적) |
| 🖼️ **이미지 추출** | 임베디드 이미지 추출 및 좌표 정보 (NEW!) |
| 🎨 **WMF/EMF 변환** | 벡터 이미지를 PNG로 변환 |
| 📷 **OCR 연동** | 외부 OCR 서비스 연동을 위한 JSON 출력 |
| 🎨 **시각화** | 문서 레이아웃을 이미지로 시각화 |
| 📋 **다양한 출력** | JSON, Markdown, 구조화된 텍스트 |
| 📑 **다중 페이지** | 여러 페이지(섹션) 자동 처리 |

---

## 📦 설치

### 필수 의존성

```bash
pip install olefile Pillow
```

### 선택적 의존성

```bash
# WMF/EMF 변환을 위해 (선택)
brew install imagemagick  # macOS
apt-get install imagemagick  # Ubuntu/Debian
```

---

## 🚀 빠른 시작

### 기본 사용법 (이미지 포함)

```python
from hwpx_parser_cursor import parse_hwpx
from hwp_parser_cursor import parse_hwp
from document_extractor import extract_document_with_images, create_visualization_report

# HWPX 파일 처리
doc = parse_hwpx("document.hwpx")

# 이미지 포함 추출
extracted = extract_document_with_images(
    doc,
    extract_images=True,           # 이미지 추출 활성화
    save_images_dir="output/images"  # 이미지 저장 디렉토리
)

# HWP 파일 처리
doc = parse_hwp("document.hwp")
extracted = extract_document_with_images(doc, extract_images=True, save_images_dir="output/images")

# 결과 출력
print(f"요소 수: {len(extracted.elements)}")
print(f"표 수: {len(extracted.tables)}")
print(f"이미지 수: {len(extracted.images)}")  # NEW!

# 전체 리포트 생성 (이미지 포함)
create_visualization_report(extracted, "output/")
```

---

## 📖 상세 사용법

### 1. 문서 파싱

```python
# HWPX (XML 기반, Open Document Format)
from hwpx_parser_cursor import parse_hwpx

doc = parse_hwpx("/path/to/document.hwpx")
print(doc.title)
print(doc.to_text())
print(doc.to_markdown())
print(doc.to_json())

# HWP (OLE Compound Document Format)
from hwp_parser_cursor import parse_hwp

doc = parse_hwp("/path/to/document.hwp")
print(doc.title)
print(doc.to_text())
print(doc.to_markdown())
print(doc.to_json())
```

### 2. 구조화된 정보 추출

```python
from document_extractor import extract_document_elements

extracted = extract_document_elements(doc)

# 기본 정보
print(f"요소 수: {len(extracted.elements)}")
print(f"제목 수: {len(extracted.headings)}")
print(f"표 수: {len(extracted.tables)}")
print(f"페이지 수: {len(extracted.pages)}")
```

### 3. 바운딩 박스 및 좌표

```python
for elem in extracted.elements:
    print(f"유형: {elem.element_type}")
    print(f"텍스트: {elem.text}")
    print(f"위치: ({elem.bbox.x}, {elem.bbox.y}) mm")
    print(f"크기: {elem.bbox.width} × {elem.bbox.height} mm")
    print(f"페이지: {elem.page + 1}")
```

### 4. 표 구조화 (LLM/RAG용)

```python
for table in extracted.tables:
    print(f"표 제목: {table.title}")
    print(f"헤더: {table.headers}")
    print(f"데이터: {table.rows}")
    
    # LLM용 구조화된 텍스트
    print(table.to_structured_text())
    # 출력:
    # [표 제목] 광고심의신청 접수정보
    # [표 헤더] 신청자 | 은행명 | 담당자명
    # [행 1] 준법감시인 | | 명칭
    
    # Markdown 테이블
    print(table.to_markdown())
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
from hwp_parser_cursor import parse_hwp
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

## 📑 페이지 처리 (단일 페이지 vs 다중 페이지)

### ✅ 다중 페이지 지원

이 파서는 **여러 페이지(섹션)를 자동으로 처리**합니다.

| 파서 | 다중 페이지 지원 | 처리 방식 |
|------|-----------------|----------|
| **HWPX** | ✅ 지원 | `section0.xml`, `section1.xml`... 자동 탐색 |
| **HWP** | ✅ 지원 | `BodyText/Section0`, `Section1`... 순차 처리 |

### 단일 페이지 처리

특정 페이지만 처리하고 싶을 때:

```python
from document_extractor import extract_document_with_images, visualize_elements

# 문서 파싱
doc = parse_hwpx("multi_page_document.hwpx")
extracted = extract_document_with_images(doc, extract_images=True)

# 페이지 수 확인
print(f"총 페이지 수: {len(extracted.pages)}")

# 특정 페이지의 요소만 필터링
page_num = 0  # 첫 번째 페이지 (0부터 시작)
page_elements = [e for e in extracted.elements if e.page == page_num]
print(f"페이지 {page_num + 1}의 요소 수: {len(page_elements)}")

# 특정 페이지의 이미지만 필터링
page_images = [img for img in extracted.images if img.page == page_num]
print(f"페이지 {page_num + 1}의 이미지 수: {len(page_images)}")

# 특정 페이지만 시각화
visualize_elements(extracted, "page_1.png", page_num=0)  # 첫 번째 페이지
visualize_elements(extracted, "page_2.png", page_num=1)  # 두 번째 페이지
```

### 모든 페이지 처리

여러 페이지를 한 번에 처리할 때:

```python
from document_extractor import extract_document_with_images, create_visualization_report

# 문서 파싱 (모든 섹션 자동 포함)
doc = parse_hwpx("multi_page_document.hwpx")
extracted = extract_document_with_images(doc, extract_images=True, save_images_dir="output/images")

# 전체 페이지 정보
print(f"총 페이지: {len(extracted.pages)}")
for page in extracted.pages:
    print(f"  페이지 {page.page_num + 1}: {page.width}mm × {page.height}mm")

# 모든 페이지 시각화 (각 페이지별 이미지 생성)
create_visualization_report(extracted, "output_dir/")
# 출력:
#   output_dir/문서명_page_001.png
#   output_dir/문서명_page_002.png
#   output_dir/문서명_page_003.png
#   output_dir/문서명_images.json (이미지 메타데이터)
#   output_dir/images/ (추출된 이미지 파일들)
#   ...
```

### 페이지별 이미지 추출

```python
# 페이지별로 이미지 그룹화
from collections import defaultdict

images_by_page = defaultdict(list)
for img in extracted.images:
    images_by_page[img.page].append(img)

# 각 페이지의 이미지 처리
for page_num in sorted(images_by_page.keys()):
    images = images_by_page[page_num]
    print(f"\n=== 페이지 {page_num + 1}의 이미지 ===")
    print(f"이미지 수: {len(images)}")
    
    for img in images:
        print(f"  - {img.filename} ({img.format.upper()})")
        print(f"    위치: ({img.bbox.x:.1f}, {img.bbox.y:.1f}) mm")
        print(f"    크기: {img.bbox.width:.1f}×{img.bbox.height:.1f} mm")
```

---

## 🎨 시각화

### 단일 페이지 시각화

```python
from document_extractor import visualize_elements

# 기본 시각화 (첫 번째 페이지)
visualize_elements(extracted, "output.png")

# 특정 페이지 지정
visualize_elements(extracted, "page2.png", page_num=1)

# 옵션 설정
visualize_elements(
    extracted,
    output_path="detailed.png",
    page_num=0,           # 표시할 페이지 (0부터)
    show_bbox=True,       # 바운딩 박스 표시
    show_text=True,       # 텍스트 표시
    show_type_colors=True,# 요소 유형별 색상
    scale=3.0,            # 확대 비율 (1mm = 3px)
    font_size=10          # 폰트 크기
)
```

### 전체 페이지 시각화

```python
from document_extractor import create_visualization_report

# 모든 페이지를 개별 이미지로 저장 + JSON + 텍스트 + 이미지
create_visualization_report(extracted, "output_dir/")

# 결과:
#   output_dir/문서명_page_001.png
#   output_dir/문서명_page_002.png
#   output_dir/문서명_extracted.json
#   output_dir/문서명_structured.txt
#   output_dir/문서명_tables.md (표가 있는 경우)
#   output_dir/문서명_images.json (이미지 메타데이터)
#   output_dir/문서명_images.md (이미지 목록)
#   output_dir/images/ (추출된 이미지 파일들)
```

### 색상 범례

| 요소 유형 | 색상 |
|----------|------|
| heading | 분홍색 (#E91E63) |
| paragraph | 파란색 (#2196F3) |
| table | 녹색 (#4CAF50) |
| table_cell | 주황색 (#FF9800) |

---

## 📁 파일 구조

```
by_cursor/
├── hwpx_parser_cursor.py    # HWPX 파서 (XML 기반)
├── hwp_parser_cursor.py     # HWP 파서 (OLE 바이너리)
├── document_extractor.py    # 통합 추출 및 시각화 모듈
├── image_extractor.py       # 이미지 추출 모듈 (NEW!)
├── demo_usage.py            # 사용 예시 스크립트
├── test_parsers.py          # 테스트 스크립트
├── README.md                # 이 파일
└── results/                 # 결과 출력 폴더
    ├── hwp_extracted/
    │   └── images/          # 추출된 이미지
    └── hwpx_extracted/
        └── images/          # 추출된 이미지
```

---

## 📊 출력 형식

### ExtractedDocument 구조

```python
ExtractedDocument(
    title="문서제목",
    source_file="경로",
    file_type="hwpx" | "hwp",
    pages=[PageInfo(...)],          # 모든 페이지 정보
    elements=[DocumentElement(...)], # 모든 요소 (page 필드로 구분)
    tables=[TableStructure(...)],
    headings=[DocumentElement(...)],
    paragraphs=[DocumentElement(...)],
    images=[ImageElement(...)],      # 추출된 이미지 (NEW!)
    metadata={...}
)
```

### ImageElement 구조

```python
ImageElement(
    image_id="BIN0001",
    filename="BIN0001.jpg",
    format="jpg",
    bbox=BoundingBox(x=20.0, y=50.0, width=150.0, height=100.0),
    page=0,                          # 이미지가 속한 페이지
    pixel_width=2481,                # 픽셀 너비
    pixel_height=3508,               # 픽셀 높이
    file_size=977927,                # 파일 크기 (bytes)
    saved_path="/path/to/image",     # 저장된 경로
    ocr_text=""                      # 외부 OCR 결과
)
```

---

## 완전한 예시: 문서 파싱부터 OCR 연동까지

```python
from pathlib import Path
from hwp_parser_cursor import parse_hwp
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

## 💡 LLM/RAG 활용 팁

### 1. 이미지 위치 기반 문맥 추출

```python
# 이미지 주변의 텍스트 추출 (이미지 설명 캡처)
for img in extracted.images:
    # 이미지와 같은 페이지의 요소들
    page_elements = [e for e in extracted.elements if e.page == img.page]
    
    # 이미지 위/아래의 텍스트 찾기
    nearby_texts = []
    for elem in page_elements:
        # 이미지 바로 위 또는 아래 50mm 이내
        if abs(elem.bbox.y - img.bbox.y) < 50:
            nearby_texts.append(elem.text)
    
    print(f"이미지 {img.filename} 주변 텍스트:")
    print("\n".join(nearby_texts))
```

### 2. 이미지와 표의 통합

```python
# 이미지와 표를 시간순으로 정렬하여 RAG 청크 생성
all_content = []

for page_num in range(len(extracted.pages)):
    page_content = {
        "page": page_num + 1,
        "text_elements": [e for e in extracted.elements if e.page == page_num],
        "tables": [t for t in extracted.tables if t.page == page_num],
        "images": [i for i in extracted.images if i.page == page_num]
    }
    all_content.append(page_content)
```

### 3. 페이지별 청킹 (RAG용)

```python
# 페이지별로 청크 생성 (이미지 포함)
chunks = []
for page_num in range(len(extracted.pages)):
    page_elements = [e for e in extracted.elements if e.page == page_num]
    page_images = [i for i in extracted.images if i.page == page_num]
    
    page_text = "\n".join(e.text for e in page_elements if e.text.strip())
    
    chunks.append({
        "page": page_num + 1,
        "text": page_text,
        "image_count": len(page_images),
        "image_files": [i.saved_path for i in page_images],
        "metadata": {
            "source": extracted.source_file,
            "page": page_num + 1,
            "total_pages": len(extracted.pages)
        }
    })
```

---

## 🧪 테스트

```bash
# 데모 실행
python demo_usage.py

# 테스트 실행
python document_extractor.py  # 전체 파이프라인 테스트
python image_extractor.py     # 이미지 추출 테스트
```

---

## 📋 요구사항

- Python 3.7+
- olefile (HWP 파싱용)
- Pillow (시각화 및 이미지 추출용)
- ImageMagick (WMF/EMF 변환용, 선택)

---

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

---

## 📝 라이센스

MIT License

---

## 🤝 기여

이슈와 PR을 환영합니다!

---

## 📚 참고 자료

- [HWPX 파일 형식 명세](https://www.hancom.com/etc/hwpDownload.do)
- [HWP 파일 형식 (OLE Compound Document)](https://github.com/hancom-io/hwpx-spec)
