# HWP/HWPX Parser for Python

한글 문서 파일(`.hwp`, `.hwpx`)을 파싱하여 텍스트, 레이아웃, 좌표 정보, 이미지를 추출하는 Python 라이브러리입니다.

**LLM/RAG 시스템에 최적화된 구조화된 출력**을 제공하며, **외부 OCR 연동**을 지원합니다.

---

## ✨ 주요 기능

| 기능 | 설명 |
|------|------|
| 📝 **텍스트 추출** | 문단, 표, 제목 등 모든 텍스트 추출 |
| 📐 **정확한 좌표** | 바운딩 박스, 절대 좌표 (mm 단위) |
| 📊 **표 구조화** | 제목/헤더/내용 분리 (LLM 친화적) |
| 🖼️ **이미지 추출** | 임베디드 이미지 추출 및 좌표 정보 |
| 🎨 **WMF/EMF 변환** | 벡터 이미지를 PNG로 변환 |
| 📷 **OCR 연동 지원** | 외부 OCR 서비스 연동용 JSON 출력 |
| 🎨 **시각화** | 문서 레이아웃을 이미지로 시각화 |
| 📑 **다중 페이지** | 여러 페이지 자동 처리 |
| 🤖 **RAG 최적화** | 청크 분할 및 메타데이터 포함 |

---

## 📦 프로젝트 구조

```
hwp/
├── by_claude/          # Claude AI 기반 파서 (권장)
│   ├── hwp_parser.py
│   ├── hwpx_parser.py
│   ├── document_extractor.py
│   ├── image_extractor.py
│   └── README.md       # 상세 사용법
│
├── by_cursor/          # Cursor AI 기반 파서
│   ├── hwp_parser_cursor.py
│   ├── hwpx_parser_cursor.py
│   ├── document_extractor.py
│   ├── image_extractor.py
│   └── README.md       # 상세 사용법
│
├── utils/              # 유틸리티
│   └── generate_diagram.py
│
├── IMAGE_PARSING_REPORT.md  # 이미지 파싱 분석 보고서
├── .gitignore
└── README.md           # 이 파일
```

---

## 🚀 빠른 시작

### 설치

```bash
# 필수 의존성
pip install olefile Pillow

# 선택적: WMF/EMF 변환 (macOS)
brew install imagemagick
```

### 기본 사용법

```python
# by_claude 사용 (권장)
from by_claude.hwp_parser import parse_hwp
from by_claude.hwpx_parser import parse_hwpx
from by_claude.document_extractor import extract_document_with_images, create_visualization_report

# 문서 파싱
doc = parse_hwp("document.hwp")  # 또는 parse_hwpx("document.hwpx")

# 이미지 포함 추출
extracted = extract_document_with_images(
    doc,
    extract_images=True,
    save_images_dir="output/images"
)

# 전체 리포트 생성
create_visualization_report(extracted, "output/")

# 결과 확인
print(f"✅ 요소: {len(extracted.elements)}개")
print(f"✅ 표: {len(extracted.tables)}개")
print(f"✅ 이미지: {len(extracted.images)}개")
```

---

## 📂 폴더별 상세 설명

### 1. `by_claude/` (권장)

Claude AI가 개발한 파서로, **더 정확한 좌표 계산**과 **개선된 페이지 처리**를 제공합니다.

**주요 특징:**
- ✅ 정확한 Y 좌표 누적 로직
- ✅ 페이지 경계 처리 개선
- ✅ 테이블 셀 병합 지원
- ✅ 이미지 GSO 좌표 추출

**파일 설명:**
- `hwp_parser.py` - HWP 파일 파싱 (OLE Compound Document)
- `hwpx_parser.py` - HWPX 파일 파싱 (XML 기반)
- `document_extractor.py` - 구조화된 정보 추출 및 시각화
- `image_extractor.py` - 이미지 추출 및 OCR 연동 지원
- `test_parsers.py` - 테스트 스크립트
- `README.md` - 상세 사용 가이드

**사용법:**
```bash
cd by_claude
python document_extractor.py  # 전체 파이프라인 테스트
python image_extractor.py     # 이미지 추출 테스트
```

[➡️ 상세 문서 보기](./by_claude/README.md)

---

### 2. `by_cursor/`

Cursor AI가 개발한 파서로, **더 많은 요소 추출**과 **상세한 메타데이터**를 제공합니다.

**주요 특징:**
- ✅ 풍부한 메타데이터
- ✅ 클래스 기반 요소 분류
- ✅ 시각화 색상 범례
- ✅ 다양한 출력 형식

**파일 설명:**
- `hwp_parser_cursor.py` - HWP 파일 파싱
- `hwpx_parser_cursor.py` - HWPX 파일 파싱
- `document_extractor.py` - 통합 추출 및 시각화
- `image_extractor.py` - 이미지 추출
- `demo_usage.py` - 사용 예시 스크립트
- `test_parsers.py` - 테스트 스크립트
- `README.md` - 상세 사용 가이드

**사용법:**
```bash
cd by_cursor
python document_extractor.py  # 전체 파이프라인 테스트
python demo_usage.py          # 데모 실행
```

[➡️ 상세 문서 보기](./by_cursor/README.md)

---

### 3. `utils/`

유틸리티 스크립트 모음

**파일 설명:**
- `generate_diagram.py` - 문서 구조 다이어그램 생성
- `hwpx_parser_diagram.md` - HWPX 파서 구조 설명

---

### 4. 기타 파일

| 파일 | 설명 |
|------|------|
| `IMAGE_PARSING_REPORT.md` | 이미지 파싱 기능 분석 및 구현 보고서 |
| `.gitignore` | Git 제외 파일 목록 |
| `pyproject.toml` | 프로젝트 의존성 (선택) |

---

## 🔄 두 파서 비교

| 기능 | by_claude | by_cursor |
|------|-----------|-----------|
| **좌표 정확도** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ |
| **페이지 처리** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ |
| **이미지 추출** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| **요소 분류** | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| **메타데이터** | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| **시각화** | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| **OCR 연동** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |

**권장 사항:**
- **정확한 좌표가 중요한 경우**: `by_claude` 사용
- **풍부한 메타데이터가 필요한 경우**: `by_cursor` 사용
- **둘 다 테스트해보고 선택하는 것을 권장합니다**

---

## 📊 지원 형식

### 문서 형식

| 형식 | 확장자 | 설명 |
|------|--------|------|
| **HWPX** | `.hwpx` | 한글 2014 이후 XML 기반 (ZIP 압축) |
| **HWP** | `.hwp` | 한글 97 이후 OLE Compound Document |

### 이미지 형식

| 형식 | 지원 | 변환 |
|------|------|------|
| JPEG/JPG | ✅ | - |
| PNG | ✅ | - |
| BMP | ✅ | zlib 압축 해제 지원 |
| GIF | ✅ | - |
| WMF | ✅ | PNG 변환 (ImageMagick 필요) |
| EMF | ✅ | PNG 변환 (ImageMagick 필요) |

---

## 🎯 사용 사례

### 1. RAG 시스템 구축
```python
# 문서 파싱 → 청크 분할 → 벡터 DB 저장
extracted = extract_document_with_images(doc, extract_images=True)
chunks = extracted.to_rag_chunks(max_chunk_size=1000)

for chunk in chunks:
    # 벡터 DB에 저장
    vector_db.add(
        text=chunk["text"],
        metadata=chunk["metadata"],
        images=[img.saved_path for img in extracted.images if img.page == chunk["page"]]
    )
```

### 2. OCR 파이프라인
```python
# 이미지 추출 → OCR → 결과 통합
extracted = extract_document_with_images(doc, extract_images=True, save_images_dir="output/images")

for img in extracted.images:
    # 외부 OCR 호출
    ocr_result = tesseract.image_to_string(img.saved_path)
    img.ocr_text = ocr_result
    
# OCR 결과 포함 JSON 저장
with open("output/document_with_ocr.json", "w") as f:
    json.dump(extracted.to_dict(), f, ensure_ascii=False, indent=2)
```

### 3. 문서 비교 및 분석
```python
# 두 문서의 구조 비교
doc1 = extract_document_with_images(parse_hwp("v1.hwp"))
doc2 = extract_document_with_images(parse_hwp("v2.hwp"))

print(f"표 변경: {len(doc1.tables)} → {len(doc2.tables)}")
print(f"이미지 변경: {len(doc1.images)} → {len(doc2.images)}")
```

---

## 🔧 의존성

### 필수
```bash
pip install olefile Pillow
```

### 선택 (이미지 변환)
```bash
# macOS
brew install imagemagick

# Ubuntu/Debian
sudo apt-get install imagemagick

# 또는 LibreOffice 사용
# macOS: brew install --cask libreoffice
```

---

## 📝 라이센스

MIT License

---

## 🤝 기여

이슈와 PR을 환영합니다!

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 📚 참고 자료

- [HWPX 파일 형식 명세](https://www.hancom.com/etc/hwpDownload.do)
- [HWP 파일 형식 (한컴 오피스)](https://github.com/hancom-io/hwpx-spec)
- [이미지 파싱 상세 분석 보고서](./IMAGE_PARSING_REPORT.md)

---

## 🎓 개발자

- **by_claude**: Claude AI (Anthropic)
- **by_cursor**: Cursor AI

---

## 📧 문의

이슈를 통해 문의해주세요.

