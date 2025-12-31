# Document Element Classes

이 문서는 HWPX/HWP 파서에서 사용하는 문서 요소 클래스 코드를 정의합니다.

## Class Codes (시각화용)

시각화 이미지에서 각 요소는 클래스 코드로 표시됩니다:

| Code | Class Name   | Description (설명)                    | Color (색상)      |
|------|--------------|---------------------------------------|-------------------|
| `h`  | heading      | 제목 (문서 구조를 나타내는 제목 텍스트)   | Pink (#E91E63)    |
| `p`  | paragraph    | 문단 (일반 본문 텍스트)                  | Blue (#2196F3)    |
| `T`  | table        | 표 (테이블 전체 영역)                    | Green (#4CAF50)   |
| `c`  | table_cell   | 표 셀 (테이블 내 개별 셀)                | Orange (#FF9800)  |
| `i`  | image        | 이미지 (그림, 사진 등)                   | Purple (#9C27B0)  |

## Layout Classes (레이아웃 클래스)

### heading (h)
- **용도**: 문서의 구조적 제목
- **판별 기준**:
  - 스타일 ID가 1-6인 경우
  - 패턴 매칭: `제N장`, `제N조`, `가.`, `1.`, `【】` 등
  - 짧고 마침표로 끝나지 않는 텍스트
- **레벨**: 1 (대제목), 2 (중제목), 3 (소제목)

### paragraph (p)
- **용도**: 일반 본문 텍스트
- **특징**: 제목이 아닌 모든 텍스트 블록

### table (T)
- **용도**: 표 전체 영역
- **속성**:
  - `rows`: 행 수
  - `cols`: 열 수
  - `bbox`: 표 전체의 바운딩 박스

### table_cell (c)
- **용도**: 표 내 개별 셀
- **속성**:
  - `row`, `col`: 셀 위치
  - `row_span`, `col_span`: 병합 정보
  - `text`: 셀 내용

### image (i)
- **용도**: 이미지, 그림, 도형
- **참고**: 현재 버전에서는 이미지 추출이 제한적임

## Text Classes (텍스트 클래스)

텍스트 레벨의 세부 분류:

| Type          | Description                              |
|---------------|------------------------------------------|
| text_run      | 동일 서식의 연속 텍스트 조각               |
| line_segment  | 한 줄의 텍스트 (레이아웃 좌표 포함)         |

## Hierarchical Structure (계층 구조)

문서는 다음과 같은 계층 구조로 구성됩니다:

```
Document
├── Section (섹션)
│   ├── Page Properties (페이지 속성)
│   └── Paragraphs (문단들)
│       ├── heading (h) - Level 1
│       │   ├── heading (h) - Level 2
│       │   │   ├── paragraph (p)
│       │   │   └── table (T)
│       │   │       └── table_cell (c)
│       │   └── heading (h) - Level 3
│       └── paragraph (p)
```

## JSON Output Structure

추출된 JSON에서 클래스 정보는 다음과 같이 표현됩니다:

```json
{
  "elements": [
    {
      "id": "elem_0001",
      "type": "heading",      // 클래스 타입
      "text": "제목 텍스트",
      "level": 1,             // 제목 레벨 (heading만 해당)
      "bbox": {
        "x": 20.0,
        "y": 10.0,
        "width": 170.0,
        "height": 6.3
      },
      "page": 0
    }
  ]
}
```

## Visualization Legend

시각화 이미지 하단에 범례가 표시됩니다:

```
Classes:
[h] = heading    (분홍)
[p] = paragraph  (파랑)
[T] = table      (초록)
[c] = cell       (주황)
```

## Usage Example

```python
from document_extractor import extract_document, visualize_elements
from hwpx_parser import parse_hwpx_file

# 문서 파싱 및 추출
doc = parse_hwpx_file("document.hwpx")
extracted = extract_document(doc)

# 클래스별 요소 필터링
headings = [e for e in extracted.elements if e.element_type == "heading"]
paragraphs = [e for e in extracted.elements if e.element_type == "paragraph"]
tables = [e for e in extracted.elements if e.element_type == "table"]

# 시각화 (클래스 코드 포함)
visualize_elements(extracted, "output.png")
```
