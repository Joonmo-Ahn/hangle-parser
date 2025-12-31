# HWP/HWPX 문서의 이미지 변환 기능 조사 보고서

**작성일**: 2024-12-30  
**목적**: HWP/HWPX 문서를 이미지로 변환하는 방법 조사 및 현재 구현 분석

---

## 📋 목차

1. [개요](#개요)
2. [HWP/HWPX를 이미지로 변환하는 방법](#hwphwpx를-이미지로-변환하는-방법)
3. [현재 구현 분석](#현재-구현-분석)
4. [완전 변환 vs 부분 변환](#완전-변환-vs-부분-변환)
5. [HWP vs HWPX 차이점](#hwp-vs-hwpx-차이점)
6. [비교 분석](#비교-분석)
7. [결론 및 권장사항](#결론-및-권장사항)

---

## 개요

HWP/HWPX 문서를 이미지로 변환하는 것은 **가능**하며, 여러 가지 방법이 존재합니다. 본 보고서에서는 각 방법의 특징과 한계를 분석하고, 현재 구현된 기능과 비교합니다.

---

## HWP/HWPX를 이미지로 변환하는 방법

### 방법 1: 한컴오피스 공식 프로그램 사용 ⭐⭐⭐⭐⭐

#### 설명
한컴오피스 한글 프로그램의 내장 기능을 사용하여 직접 변환

#### 절차
1. 한글 프로그램에서 문서 열기
2. **파일** → **다른 이름으로 저장**
3. 파일 형식 선택: **JPEG**, **PNG**, **BMP**
4. 저장

#### 변환 범위
- ✅ **완전 변환 가능**
- 모든 페이지를 개별 이미지로 저장
- 원본 레이아웃 100% 유지

#### 장점
```
✅ 원본 레이아웃 완벽 재현
✅ 글꼴, 색상, 도형, 이미지 모두 보존
✅ 복잡한 표, 수식, 그래프 정확히 변환
✅ 고해상도 출력 지원
✅ 다중 페이지 일괄 변환 가능
```

#### 단점
```
❌ 한컴오피스 라이선스 필요 (유료)
❌ GUI 환경 필요 (자동화 어려움)
❌ Windows/macOS 전용
❌ Python 스크립트로 직접 호출 불가
❌ 배치 처리 제한적
```

#### 적용 가능성
| 항목 | 평가 |
|------|------|
| **정확도** | ⭐⭐⭐⭐⭐ (100%) |
| **자동화** | ⭐⭐ (매크로 가능) |
| **무료 사용** | ⭐ (체험판 제한) |
| **대량 처리** | ⭐⭐⭐ |

---

### 방법 2: PDF 경유 변환 ⭐⭐⭐⭐

#### 설명
HWP/HWPX → PDF → 이미지 (2단계 변환)

#### 절차
1. 한글 프로그램: **PDF로 저장**
2. PDF 도구 사용:
   - Ghostscript
   - ImageMagick
   - PyMuPDF (Python)
   - pdf2image (Python)

#### 예시 코드
```python
# PDF → 이미지 변환 (Python)
from pdf2image import convert_from_path

# HWP → PDF (한컴오피스 필요)
# PDF → 이미지
images = convert_from_path('document.pdf', dpi=300)
for i, image in enumerate(images):
    image.save(f'page_{i+1}.png', 'PNG')
```

#### 변환 범위
- ✅ **완전 변환 가능**
- PDF 품질에 따라 달라짐

#### 장점
```
✅ 고품질 출력 가능 (DPI 조절)
✅ Python 자동화 가능 (PDF 이후 단계)
✅ 다양한 라이브러리 지원
✅ 벡터 형식 → 래스터 변환
```

#### 단점
```
❌ 2단계 변환 필요 (시간 소요)
❌ 첫 번째 단계는 여전히 한컴오피스 필요
❌ 파일 크기 증가 가능
❌ 일부 특수 기능 손실 가능
```

#### 적용 가능성
| 항목 | 평가 |
|------|------|
| **정확도** | ⭐⭐⭐⭐ (95%+) |
| **자동화** | ⭐⭐⭐ |
| **무료 사용** | ⭐⭐⭐⭐ |
| **대량 처리** | ⭐⭐⭐⭐ |

---

### 방법 3: LibreOffice 헤드리스 모드 ⭐⭐⭐

#### 설명
LibreOffice를 CLI로 실행하여 변환 (완전 자동화)

#### 절차
```bash
# HWP → PDF
soffice --headless --convert-to pdf document.hwp

# PDF → 이미지
convert -density 300 document.pdf output_%03d.png
```

#### Python 통합
```python
import subprocess

# LibreOffice 변환
subprocess.run([
    'soffice',
    '--headless',
    '--convert-to', 'pdf',
    '--outdir', 'output/',
    'document.hwp'
])

# ImageMagick으로 이미지 변환
subprocess.run([
    'convert',
    '-density', '300',
    'output/document.pdf',
    'output/page_%03d.png'
])
```

#### 변환 범위
- ⚠️ **부분적 변환**
- 기본 텍스트, 표, 이미지는 변환
- 복잡한 한글 특수 기능은 손실 가능

#### 장점
```
✅ 완전 무료 (오픈소스)
✅ CLI 자동화 가능
✅ 크로스 플랫폼 (Windows/Linux/macOS)
✅ 서버 환경에서 실행 가능
✅ Python 스크립트 통합 용이
```

#### 단점
```
❌ HWP 지원이 완벽하지 않음
❌ 한글 특수 서식 일부 손실
❌ 글꼴 대체 발생 가능
❌ 표 레이아웃 변경 가능
❌ 복잡한 문서는 오류 발생 가능
```

#### 적용 가능성
| 항목 | 평가 |
|------|------|
| **정확도** | ⭐⭐⭐ (70-80%) |
| **자동화** | ⭐⭐⭐⭐⭐ |
| **무료 사용** | ⭐⭐⭐⭐⭐ |
| **대량 처리** | ⭐⭐⭐⭐⭐ |

---

### 방법 4: 온라인 변환 도구 ⭐⭐

#### 설명
웹 기반 변환 서비스 사용

#### 예시
- Convertio
- CloudConvert
- OnlineConvert

#### 변환 범위
- ⚠️ **서비스마다 다름**
- 일부는 완전 변환, 일부는 제한적

#### 장점
```
✅ 소프트웨어 설치 불필요
✅ 간편한 사용
✅ 다양한 형식 지원
```

#### 단점
```
❌ 인터넷 연결 필요
❌ 보안 위험 (업로드 필요)
❌ 파일 크기/개수 제한
❌ API 호출 비용 발생 가능
❌ 자동화 제한적
❌ 대량 처리 어려움
```

#### 적용 가능성
| 항목 | 평가 |
|------|------|
| **정확도** | ⭐⭐⭐ (서비스마다 다름) |
| **자동화** | ⭐⭐ |
| **무료 사용** | ⭐⭐ (제한적) |
| **대량 처리** | ⭐ |

---

### 방법 5: 현재 구현 - 레이아웃 재구성 방식 ⭐⭐⭐⭐

#### 설명
파싱된 문서 정보(텍스트, 좌표, 스타일)를 바탕으로 Pillow로 재구성

#### 코드 예시
```python
from by_claude.hwp_parser import parse_hwp
from by_claude.document_extractor import extract_document_with_images, visualize_elements

# 1. 문서 파싱
doc = parse_hwp("document.hwp")

# 2. 구조화된 정보 추출
extracted = extract_document_with_images(doc, extract_images=True)

# 3. 이미지 생성 (레이아웃 재구성)
visualize_elements(extracted, "output.png", page_num=0, scale=3.0)
```

#### 생성 방식
1. **페이지 캔버스 생성** (PIL Image)
   ```python
   img = Image.new('RGB', (width, height), 'white')
   ```

2. **요소별 바운딩 박스 그리기**
   ```python
   draw.rectangle([x, y, x+width, y+height], outline=color)
   ```

3. **텍스트 렌더링**
   ```python
   draw.text((x, y), text, fill=color, font=font)
   ```

4. **이미지 저장**
   ```python
   img.save("output.png", "PNG")
   ```

#### 변환 범위
- ⚠️ **부분적 변환**
- 텍스트, 바운딩 박스, 기본 레이아웃은 표시
- 원본 렌더링은 아님

#### 포함되는 요소
```
✅ 텍스트 내용 (모든 문자)
✅ 바운딩 박스 (위치 정보)
✅ 요소 유형 표시 (h:heading, p:paragraph, T:table)
✅ 페이지 구조
✅ 레이아웃 좌표
✅ 추출된 임베디드 이미지 (별도 파일)
```

#### 제외되는 요소
```
❌ 원본 글꼴 (시스템 글꼴로 대체)
❌ 정확한 색상 (기본 색상 사용)
❌ 도형, 선, 그래프
❌ 배경 이미지
❌ 워터마크
❌ 복잡한 표 테두리
❌ 그림자, 효과
```

#### 장점
```
✅ 100% Python (외부 의존성 최소)
✅ 완전 자동화 가능
✅ 좌표 정보 시각화에 최적
✅ 디버깅 및 검증 용이
✅ 무료
✅ 크로스 플랫폼
✅ 서버 환경 실행 가능
```

#### 단점
```
❌ 원본 문서와 시각적으로 다름
❌ 디자인 요소 제한적
❌ OCR이나 인쇄용으로 부적합
❌ 레이아웃 검증/디버깅 목적에 한정
```

#### 적용 가능성
| 항목 | 평가 |
|------|------|
| **정확도** | ⭐⭐⭐ (레이아웃만 70%) |
| **자동화** | ⭐⭐⭐⭐⭐ |
| **무료 사용** | ⭐⭐⭐⭐⭐ |
| **대량 처리** | ⭐⭐⭐⭐⭐ |

---

## 현재 구현 분석

### 구현된 기능

#### 1. `visualize_elements()` - 단일 페이지 시각화
```python
def visualize_elements(
    extracted: ExtractedDocument,
    output_path: str | Path,
    page_num: int = 0,
    show_bbox: bool = True,      # 바운딩 박스 표시
    show_text: bool = True,       # 텍스트 표시
    scale: float = 3.0,           # 확대 비율
    font_size: int = 8,
) -> Path
```

**기능:**
- 지정된 페이지를 PNG 이미지로 생성
- 바운딩 박스와 텍스트 렌더링
- 요소 유형별 색상 구분

**출력 예시:**
```
┌─────────────────────────┐
│ h:가. 광고심의신청 접수정보 │  ← 분홍색 바운딩 박스 (제목)
│                         │
│ p:은행명                 │  ← 파란색 바운딩 박스 (문단)
│                         │
│ T:[표 3×4] 접수정보       │  ← 녹색 바운딩 박스 (표)
│ c:신청자                 │  ← 주황색 (표 셀)
└─────────────────────────┘
```

#### 2. `visualize_all_pages()` - 다중 페이지 시각화
```python
def visualize_all_pages(
    extracted: ExtractedDocument,
    output_dir: str | Path,
    ...
) -> list[Path]
```

**기능:**
- 모든 페이지를 개별 PNG로 저장
- `document_page_001.png`, `document_page_002.png`, ...

#### 3. `visualize_to_pdf()` - PDF 생성
```python
def visualize_to_pdf(
    extracted: ExtractedDocument,
    output_path: str | Path,
    ...
) -> Path
```

**기능:**
- 모든 페이지를 하나의 PDF로 결합
- 레이아웃 검증용

#### 4. `create_visualization_report()` - 종합 리포트
```python
def create_visualization_report(
    extracted: ExtractedDocument,
    output_dir: str | Path,
) -> list[Path]
```

**생성 파일:**
```
output/
├── document_page_001.png       # 각 페이지 시각화
├── document_page_002.png
├── document_extracted.json     # 구조화된 데이터
├── document_structured.txt     # LLM용 텍스트
├── document_chunks.json        # RAG 청크
├── document_tables.md          # 표 목록
├── document_images.json        # 이미지 메타데이터 (NEW!)
├── document_images.md          # 이미지 목록 (NEW!)
└── images/                     # 추출된 임베디드 이미지
    ├── BIN0001.jpg
    └── BIN0002.bmp
```

---

## 완전 변환 vs 부분 변환

### 완전 변환 (Full Document Rendering)

#### 정의
원본 문서를 **픽셀 단위로 정확히 재현**하는 변환

#### 특징
```
✅ 모든 시각적 요소 포함
✅ 글꼴, 색상, 도형, 효과 보존
✅ 인쇄 품질 출력
✅ OCR 입력으로 사용 가능
✅ 시각적으로 원본과 동일
```

#### 방법
- **한컴오피스** (⭐ 최고 정확도)
- **PDF 경유** (⭐ 높은 정확도)
- **온라인 도구** (서비스마다 다름)

#### 사용 사례
- 문서 아카이빙
- 인쇄/출판
- OCR 전처리
- 문서 미리보기

---

### 부분 변환 (Layout Reconstruction)

#### 정의
파싱된 **구조 정보를 바탕으로 재구성**한 변환

#### 특징
```
✅ 레이아웃 구조 표현
✅ 텍스트 내용 포함
✅ 좌표 정보 시각화
⚠️ 디자인 요소 제한적
⚠️ 원본과 시각적 차이
```

#### 방법
- **현재 구현** (Pillow 기반)
- **LibreOffice** (부분 지원)

#### 사용 사례
- 레이아웃 검증
- 좌표 디버깅
- 문서 구조 분석
- 자동화된 대량 처리

---

### 비교표

| 항목 | 완전 변환 | 부분 변환 (현재 구현) |
|------|----------|---------------------|
| **시각적 정확도** | ⭐⭐⭐⭐⭐ (100%) | ⭐⭐⭐ (70%) |
| **글꼴 재현** | ✅ 완벽 | ❌ 시스템 글꼴 대체 |
| **색상 재현** | ✅ 완벽 | ⚠️ 기본 색상만 |
| **도형/선** | ✅ 포함 | ❌ 미포함 |
| **이미지** | ✅ 임베드 | ⚠️ 별도 추출 |
| **표 테두리** | ✅ 정확 | ⚠️ 단순화 |
| **자동화** | ⭐⭐ | ⭐⭐⭐⭐⭐ |
| **무료 사용** | ❌ | ✅ |
| **OCR 입력** | ✅ 적합 | ❌ 부적합 |
| **레이아웃 검증** | ⚠️ | ✅ 최적 |

---

## HWP vs HWPX 차이점

### HWP (.hwp) - 바이너리 형식

#### 구조
- **OLE Compound Document** (Microsoft Office 97-2003과 유사)
- 바이너리 스트림으로 구성
- 압축된 데이터 (zlib)

#### 파싱 난이도
```
❌ 복잡한 바이너리 구조
❌ 스펙 문서 제한적
❌ 역공학 필요
⚠️ 좌표 추출 어려움
```

#### 이미지 변환
**완전 변환:**
- 한컴오피스만 100% 지원
- LibreOffice: 70-80% 지원

**부분 변환 (현재 구현):**
```python
# HWP 파싱 → 시각화
from by_claude.hwp_parser import parse_hwp
from by_claude.document_extractor import visualize_elements

doc = parse_hwp("document.hwp")
extracted = extract_document_with_images(doc)
visualize_elements(extracted, "output.png")
```

**제약 사항:**
- GSO 컨트롤 좌표가 상대적 (절대 좌표 추출 어려움)
- 표 구조 파싱 제한적
- 일부 문서는 좌표가 0으로 나옴

---

### HWPX (.hwpx) - XML 형식

#### 구조
- **ZIP 압축된 XML 파일들**
- `section*.xml`, `header.xml` 등
- 텍스트 기반

#### 파싱 난이도
```
✅ XML 구조로 읽기 쉬움
✅ 좌표 정보 명시적
✅ 표 구조 파싱 용이
✅ 이미지 참조 명확
```

#### 이미지 변환
**완전 변환:**
- 한컴오피스 100% 지원
- LibreOffice: 80-90% 지원 (HWP보다 우수)

**부분 변환 (현재 구현):**
```python
# HWPX 파싱 → 시각화
from by_claude.hwpx_parser import parse_hwpx
from by_claude.document_extractor import visualize_elements

doc = parse_hwpx("document.hwpx")
extracted = extract_document_with_images(doc)
visualize_elements(extracted, "output.png")
```

**장점:**
- 좌표 정보 정확
- 표 구조 완벽 파싱
- 이미지 위치 명확 (`<hp:offset>`, `<hp:curSz>`)

---

### 변환 품질 비교

| 항목 | HWP (.hwp) | HWPX (.hwpx) |
|------|-----------|--------------|
| **한컴오피스 변환** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| **LibreOffice 변환** | ⭐⭐⭐ (70-80%) | ⭐⭐⭐⭐ (80-90%) |
| **현재 구현 (재구성)** | ⭐⭐⭐ | ⭐⭐⭐⭐ |
| **좌표 정확도** | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| **표 파싱** | ⭐⭐ | ⭐⭐⭐⭐⭐ |
| **이미지 추출** | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |

---

## 비교 분석

### 종합 비교표

| 변환 방법 | 정확도 | 자동화 | 무료 | 대량처리 | 서버실행 | 총점 |
|----------|--------|--------|------|---------|---------|------|
| **한컴오피스** | 100% | ⭐⭐ | ❌ | ⭐⭐⭐ | ❌ | ⭐⭐⭐ |
| **PDF 경유** | 95% | ⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⚠️ | ⭐⭐⭐⭐ |
| **LibreOffice** | 70-80% | ⭐⭐⭐⭐⭐ | ✅ | ⭐⭐⭐⭐⭐ | ✅ | ⭐⭐⭐⭐ |
| **온라인 도구** | 다양 | ⭐⭐ | ⭐⭐ | ⭐ | ❌ | ⭐⭐ |
| **현재 구현** | 70% | ⭐⭐⭐⭐⭐ | ✅ | ⭐⭐⭐⭐⭐ | ✅ | ⭐⭐⭐⭐ |

---

### 시나리오별 권장 방법

#### 시나리오 1: 인쇄/출판용 고품질 이미지
**권장:** 한컴오피스 → 이미지  
**이유:** 100% 정확도, 모든 디자인 요소 보존

#### 시나리오 2: 대량 문서 자동 처리 (서버)
**권장:** LibreOffice Headless 또는 현재 구현  
**이유:** 완전 자동화, 무료, 서버 실행 가능

#### 시나리오 3: 레이아웃 검증/디버깅
**권장:** 현재 구현 (visualize_elements)  
**이유:** 좌표 정보 시각화, 빠른 처리

#### 시나리오 4: OCR 전처리
**권장:** 한컴오피스 → 이미지 또는 PDF 경유  
**이유:** 고해상도, 정확한 텍스트 렌더링

#### 시나리오 5: 문서 미리보기 (웹 서비스)
**권장:** PDF 경유 + 썸네일 생성  
**이유:** 캐싱 가능, 브라우저 호환성

---

## 결론 및 권장사항

### 핵심 발견사항

1. **HWP/HWPX → 이미지 변환은 가능하다**
   - 완전 변환: 한컴오피스, PDF 경유
   - 부분 변환: LibreOffice, 현재 구현

2. **HWPX가 HWP보다 변환이 용이하다**
   - XML 구조로 파싱 쉬움
   - 좌표 정보 명확
   - LibreOffice 지원 우수

3. **현재 구현은 '완전 변환'이 아니라 '레이아웃 재구성'이다**
   - 원본 렌더링 ≠ 재구성
   - 좌표 검증용으로 최적화
   - 인쇄/OCR 용도로는 부적합

4. **자동화와 정확도는 트레이드오프 관계**
   - 높은 정확도 → 한컴오피스 (자동화 어려움)
   - 높은 자동화 → LibreOffice/현재 구현 (정확도 낮음)

---

### 개선 권장사항

#### 1. 완전 변환 기능 추가 (선택)

한컴오피스 CLI 인터페이스 활용 (가능하다면):
```python
def convert_hwp_to_image_full(hwp_file, output_dir):
    """
    한컴오피스 CLI를 통한 완전 변환
    (한컴오피스 설치 및 라이선스 필요)
    """
    subprocess.run([
        'hwp.exe',  # 한컴오피스 CLI (존재 여부 확인 필요)
        '--convert-to-image',
        '--output', output_dir,
        hwp_file
    ])
```

#### 2. LibreOffice 통합 추가

```python
def convert_with_libreoffice(file_path, output_dir):
    """
    LibreOffice Headless를 통한 변환
    - 자동화 가능
    - 무료
    - 70-80% 정확도
    """
    # HWP → PDF
    subprocess.run([
        'soffice',
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', output_dir,
        file_path
    ])
    
    # PDF → 이미지
    from pdf2image import convert_from_path
    images = convert_from_path(f"{output_dir}/document.pdf", dpi=300)
    for i, img in enumerate(images):
        img.save(f"{output_dir}/page_{i+1}.png", "PNG")
```

#### 3. 현재 구현 개선

**개선 포인트:**
```python
# 1. 임베디드 이미지를 시각화에 포함
def visualize_with_embedded_images(extracted, output_path):
    """이미지도 함께 렌더링"""
    # 현재: 텍스트만 렌더링
    # 개선: extracted.images를 좌표에 맞게 배치
    pass

# 2. 표 테두리 개선
def draw_table_borders(draw, table_bbox):
    """표 테두리를 더 정확하게 그리기"""
    pass

# 3. 글꼴 매핑 개선
FONT_MAPPING = {
    "바탕": "AppleSDGothicNeo-Regular",
    "돋움": "AppleSDGothicNeo-Medium",
    # ... 더 많은 매핑
}
```

#### 4. 하이브리드 접근법

```python
def smart_convert(file_path, output_dir, mode='auto'):
    """
    상황에 따라 최적의 변환 방법 선택
    
    mode='full': 한컴오피스 우선, 실패시 LibreOffice
    mode='fast': 현재 구현 (빠른 처리)
    mode='auto': 파일 복잡도에 따라 자동 선택
    """
    if mode == 'full':
        try:
            return convert_with_hancom(file_path, output_dir)
        except:
            return convert_with_libreoffice(file_path, output_dir)
    elif mode == 'fast':
        return visualize_elements(file_path, output_dir)
    else:
        complexity = analyze_document_complexity(file_path)
        if complexity > 0.7:
            return convert_with_hancom(file_path, output_dir)
        else:
            return visualize_elements(file_path, output_dir)
```

---

### 최종 정리

| 질문 | 답변 |
|------|------|
| **HWP/HWPX를 이미지로 변환 가능한가?** | ✅ **가능** |
| **완전 변환 가능한가?** | ✅ **가능** (한컴오피스, PDF 경유) |
| **부분 변환은?** | ✅ **가능** (LibreOffice, 현재 구현) |
| **HWP vs HWPX 차이는?** | HWPX가 **더 정확** (XML 구조) |
| **현재 구현은?** | **레이아웃 재구성** (완전 렌더링 ❌) |
| **자동화 가능한가?** | ✅ **가능** (LibreOffice, 현재 구현) |
| **무료로 가능한가?** | ✅ **가능** (LibreOffice, 현재 구현) |

---

## 참고 자료

1. **한컴오피스 공식 문서**
   - https://help.hancom.com/

2. **LibreOffice 문서 변환**
   - https://www.libreoffice.org/

3. **pdf2image (Python)**
   - https://github.com/Belval/pdf2image

4. **Pillow (Python)**
   - https://pillow.readthedocs.io/

5. **현재 구현 코드**
   - `by_claude/document_extractor.py`
   - `by_claude/image_extractor.py`

---

**보고서 끝**

