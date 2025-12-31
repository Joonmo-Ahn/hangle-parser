# 📷 HWP/HWPX 이미지 파싱 기술 보고서

## 📌 요약

| 파일 형식 | 이미지 포함 여부 | 현재 파싱 지원 | 기술적 구현 난이도 |
|----------|-----------------|---------------|------------------|
| **HWP** | ✅ 지원 (BinData/) | ❌ 미지원 | ⭐⭐⭐ 중간 |
| **HWPX** | ✅ 지원 (BinData/) | ❌ 미지원 | ⭐⭐ 쉬움 |

---

## 1. HWP 파일 이미지 구조 분석

### 1.1 OLE 스트림 구조

```
HWP 파일 (OLE Compound Document)
├── FileHeader          # 파일 헤더
├── DocInfo             # 문서 정보
├── BodyText/
│   └── Section0        # 본문 (이미지 참조 포함)
├── BinData/            # ✨ 바이너리 데이터 (이미지)
│   ├── BIN0001.jpg     # JPEG 이미지 (977KB)
│   ├── BIN0002.bmp     # BMP 이미지 (zlib 압축)
│   └── ...
├── PrvImage            # 미리보기 이미지
└── Scripts/            # 스크립트
```

### 1.2 테스트 파일 분석: `2. [농협] 광고안(B).hwp`

```
발견된 이미지:
┌─────────────────┬───────────────┬─────────────────┐
│ 파일명          │ 크기          │ 형식            │
├─────────────────┼───────────────┼─────────────────┤
│ BIN0001.jpg     │ 977,927 bytes │ JPEG (비압축)   │
│ BIN0002.bmp     │ 2,042 bytes   │ BMP (zlib 압축) │
│ (해제 후)       │ 14,958 bytes  │ BMP 원본        │
└─────────────────┴───────────────┴─────────────────┘
```

### 1.3 이미지 참조 메커니즘

HWP 파일에서 이미지는 **CTRL_HEADER** 레코드의 **GSO (Graphic Shape Object)** 컨트롤을 통해 참조됩니다:

```
BodyText/Section0 레코드 구조:
┌────────────────────────────────────────────────────────┐
│ CTRL_HEADER (태그: 0x47)                               │
│   - 컨트롤 ID: " osg" (gso 역순 = Graphic Shape Object)│
│   - 위치: x, y (HWPUNIT)                               │
│   - 크기: width, height (HWPUNIT)                      │
│   - BinData 참조 ID: BIN0001.jpg → ID 0                │
└────────────────────────────────────────────────────────┘
```

### 1.4 이미지 압축 방식

| 확장자 | 저장 방식 | 처리 방법 |
|-------|---------|----------|
| `.jpg`, `.jpeg` | 비압축 (원본) | 직접 읽기 |
| `.png` | 비압축 (원본) | 직접 읽기 |
| `.bmp` | zlib 압축 | `zlib.decompress(data, -15)` |
| `.gif` | 비압축 (원본) | 직접 읽기 |
| `.wmf`, `.emf` | zlib 압축 | 압축 해제 후 변환 필요 |

---

## 2. HWPX 파일 이미지 구조 분석

### 2.1 ZIP 아카이브 구조

```
HWPX 파일 (ZIP 아카이브)
├── mimetype
├── version.xml
├── Contents/
│   ├── header.xml      # 헤더 (이미지 참조 정보)
│   ├── section0.xml    # 본문 (이미지 태그 포함)
│   └── content.hpf
├── BinData/            # ✨ 바이너리 데이터 (이미지)
│   ├── image1.jpg
│   ├── image2.png
│   └── ...
└── Preview/
    └── PrvImage.png    # 미리보기 이미지
```

### 2.2 테스트 파일 분석: `은행권 광고심의 결과 보고서(양식)vF (1).hwpx`

```
발견된 파일 목록:
- mimetype
- version.xml
- Contents/header.xml (163KB)
- Contents/section0.xml (500KB)
- Preview/PrvImage.png (24KB) ← 미리보기 이미지만 존재
- settings.xml
- META-INF/...

→ BinData/ 폴더 없음 (본문에 이미지 미포함)
```

### 2.3 이미지 참조 방식 (XML)

```xml
<!-- section0.xml 내 이미지 참조 예시 -->
<hp:pic>
  <hp:shapeObject>
    <hp:offset x="1000" y="2000"/>
    <hp:orgSz width="5000" height="3000"/>
    <hp:curSz width="5000" height="3000"/>
  </hp:shapeObject>
  <hp:imgDim dimwidth="100" dimheight="75"/>
  <hp:imgData mode="id" binary="BIN0001.jpg"/>
</hp:pic>
```

---

## 3. 현재 파서 상태 분석

### 3.1 by_claude 파서

| 컴포넌트 | 이미지 처리 상태 |
|---------|----------------|
| `hwp_parser.py` | ❌ BinData 스트림 읽기 없음 |
| `hwpx_parser.py` | ❌ BinData 폴더 처리 없음 |
| `document_extractor.py` | ❌ 이미지 요소 타입 없음 |

### 3.2 by_cursor 파서

| 컴포넌트 | 이미지 처리 상태 |
|---------|----------------|
| `hwp_parser_cursor.py` | ❌ CTRL_HEADER 파싱 없음 |
| `hwpx_parser_cursor.py` | ❌ `<hp:pic>` 태그 처리 없음 |
| `document_extractor.py` | ❌ 이미지 요소 타입 없음 |

---

## 4. 구현 방안

### 4.1 HWP 이미지 추출 구현

```python
# 추가 필요한 데이터 클래스
@dataclass
class EmbeddedImage:
    """임베디드 이미지"""
    bin_id: str          # BIN0001
    filename: str        # BIN0001.jpg
    format: str          # jpg, png, bmp
    data: bytes          # 원본 바이너리 데이터
    size: tuple[int, int]  # (width, height) pixels
    
    # 위치 정보 (GSO에서 추출)
    x: float = 0.0       # mm
    y: float = 0.0       # mm
    width: float = 0.0   # mm
    height: float = 0.0  # mm
    page: int = 0        # 페이지 번호

# HWP 이미지 추출 함수
def extract_images_from_hwp(hwp_file: str) -> list[EmbeddedImage]:
    """
    HWP 파일에서 모든 이미지 추출
    
    구현 단계:
    1. OLE 파일 열기
    2. BinData/ 스트림 순회
    3. 각 이미지 파일 읽기 (압축 해제)
    4. BodyText에서 GSO 컨트롤 파싱
    5. 이미지-위치 매핑
    """
    pass
```

### 4.2 HWPX 이미지 추출 구현

```python
def extract_images_from_hwpx(hwpx_file: str) -> list[EmbeddedImage]:
    """
    HWPX 파일에서 모든 이미지 추출
    
    구현 단계:
    1. ZIP 파일 열기
    2. BinData/ 폴더 순회
    3. 각 이미지 파일 읽기
    4. section*.xml에서 <hp:pic> 태그 파싱
    5. 이미지-위치 매핑
    """
    pass
```

### 4.3 필요한 수정 사항

| 파일 | 수정 내용 |
|-----|---------|
| `hwp_parser.py` | `_parse_bindata()`, `_parse_gso_control()` 추가 |
| `hwpx_parser.py` | `_parse_images()`, `_parse_pic_element()` 추가 |
| `document_extractor.py` | `DocumentElement.type = "image"` 지원 |

---

## 5. 예상 작업량

### 5.1 HWP 파서 (중간 난이도)

| 작업 | 예상 시간 | 난이도 |
|-----|---------|-------|
| BinData 스트림 추출 | 1시간 | ⭐⭐ |
| GSO 컨트롤 파싱 | 3시간 | ⭐⭐⭐ |
| 이미지-위치 매핑 | 2시간 | ⭐⭐⭐ |
| 테스트 및 디버깅 | 2시간 | ⭐⭐ |
| **합계** | **8시간** | |

### 5.2 HWPX 파서 (쉬움)

| 작업 | 예상 시간 | 난이도 |
|-----|---------|-------|
| BinData 폴더 추출 | 30분 | ⭐ |
| XML 태그 파싱 | 1시간 | ⭐⭐ |
| 이미지-위치 매핑 | 1시간 | ⭐⭐ |
| 테스트 및 디버깅 | 1시간 | ⭐ |
| **합계** | **3.5시간** | |

---

## 6. 제한 사항 및 고려 사항

### 6.1 지원 가능한 이미지 형식

| 형식 | 지원 | 비고 |
|-----|-----|------|
| JPEG | ✅ | 직접 지원 |
| PNG | ✅ | 직접 지원 |
| BMP | ✅ | zlib 압축 해제 필요 |
| GIF | ✅ | 직접 지원 |
| WMF/EMF | ⚠️ | 벡터 형식, 변환 필요 |
| OLE 객체 | ⚠️ | 복잡한 파싱 필요 |

### 6.2 위치 정보 정확도

- **HWP**: GSO 컨트롤의 위치 정보가 HWPUNIT 단위로 저장됨
- **HWPX**: XML에 명시적 좌표 포함 (더 정확함)
- 두 형식 모두 **앵커 타입** (문단/페이지/글자) 고려 필요

### 6.3 OLE 임베디드 객체

- Excel 차트, Word 문서 등 OLE 객체는 별도 처리 필요
- 현재 스코프에서 제외 권장

---

## 7. 결론 및 권장 사항

### ✅ 즉시 구현 가능
1. **HWPX 이미지 추출**: ZIP 구조로 구현 간단
2. **HWP 기본 이미지 추출**: JPG/PNG 직접 읽기

### ⚠️ 추가 연구 필요
1. **HWP GSO 컨트롤 파싱**: 위치 정보 정확도 검증 필요
2. **WMF/EMF 변환**: 외부 라이브러리 의존성

### 📌 권장 구현 순서
1. HWPX BinData 폴더 이미지 추출
2. HWP BinData 스트림 이미지 추출
3. 이미지 위치 정보 추출 (GSO/XML)
4. document_extractor에 이미지 타입 통합

---

*보고서 작성일: 2025년 12월 30일*

