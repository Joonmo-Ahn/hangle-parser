# HWPX Parser 코드 구조 다이어그램

이 문서는 `hwpx_folder_parser_cursor.py`의 클래스와 함수 구조를 시각화합니다.

---

## 1. 클래스 계층 구조 (Class Hierarchy)

```mermaid
classDiagram
    direction TB
    
    %% 레이아웃 관련 클래스
    class Position {
        +str vert_rel_to
        +str horz_rel_to
        +str vert_align
        +str horz_align
        +int vert_offset
        +int horz_offset
        +bool treat_as_char
        +bool flow_with_text
        +to_mm() dict
    }
    
    class Size {
        +int width
        +int height
        +str width_rel_to
        +str height_rel_to
        +to_mm() dict
    }
    
    class Margin {
        +int left
        +int right
        +int top
        +int bottom
        +to_mm() dict
    }
    
    class LineSegment {
        +int text_pos
        +int vert_pos
        +int vert_size
        +int text_height
        +int baseline
        +int spacing
        +int horz_pos
        +int horz_size
        +to_mm() dict
    }
    
    class PageProperties {
        +int width
        +int height
        +str landscape
        +str gutter_type
        +Margin margin
        +to_mm() dict
    }
    
    %% 콘텐츠 관련 클래스
    class TableCell {
        +int row
        +int col
        +str text
        +int row_span
        +int col_span
        +Size size
        +Margin margin
        +str border_fill_id
    }
    
    class Table {
        +int rows
        +int cols
        +list~TableCell~ cells
        +str id
        +int z_order
        +Position position
        +Size size
        +Margin outer_margin
        +Margin inner_margin
        +to_markdown() str
        +to_markdown_with_layout() str
    }
    
    class TextRun {
        +str text
        +str char_pr_id
    }
    
    class Paragraph {
        +str id
        +list~str~ texts
        +list~TextRun~ text_runs
        +list~Table~ tables
        +str para_pr_id
        +str style_id
        +list~LineSegment~ line_segments
        +bool page_break
        +bool column_break
        +full_text() str
        +get_bounding_box() dict
    }
    
    class Section {
        +int index
        +list~Paragraph~ paragraphs
        +PageProperties page_props
        +full_text() str
    }
    
    class VersionInfo {
        +str application
        +str app_version
        +str xml_version
    }
    
    class HwpxDocument {
        +Path folder_path
        +VersionInfo version
        +list~Section~ sections
        +dict metadata
        +title() str
        +to_text() str
        +to_markdown() str
        +to_markdown_with_layout() str
        +to_json() str
        +to_json_with_layout() str
    }
    
    class HwpxFolderParser {
        +Path folder_path
        +Path contents_dir
        +parse() HwpxDocument
        -_parse_version() VersionInfo
        -_parse_metadata() dict
        -_parse_sections() Iterator~Section~
        -_parse_section() Section
        -_parse_page_properties() PageProperties
        -_parse_paragraph() Paragraph
        -_parse_table() Table
        -_parse_table_cell() TableCell
        -_strip_ns() str
    }
    
    %% 관계 정의
    PageProperties *-- Margin : contains
    
    TableCell *-- Size : contains
    TableCell *-- Margin : contains
    
    Table *-- TableCell : contains
    Table *-- Position : contains
    Table *-- Size : contains
    Table *-- Margin : outer_margin
    Table *-- Margin : inner_margin
    
    Paragraph *-- TextRun : contains
    Paragraph *-- Table : contains
    Paragraph *-- LineSegment : contains
    
    Section *-- Paragraph : contains
    Section *-- PageProperties : contains
    
    HwpxDocument *-- VersionInfo : contains
    HwpxDocument *-- Section : contains
    
    HwpxFolderParser ..> HwpxDocument : creates
```

---

## 2. 파싱 플로우 (Parsing Flow)

```mermaid
flowchart TD
    subgraph 입력["📁 입력"]
        A[HWPX 폴더]
        A1[version.xml]
        A2[Contents/header.xml]
        A3[Contents/section0.xml]
    end
    
    subgraph 파서["🔧 HwpxFolderParser"]
        B[parse]
        B1[_parse_version]
        B2[_parse_metadata]
        B3[_parse_sections]
        B4[_parse_section]
        B5[_parse_page_properties]
        B6[_parse_paragraph]
        B7[_parse_table]
        B8[_parse_table_cell]
    end
    
    subgraph 출력["📄 출력"]
        C[HwpxDocument]
        C1[to_text]
        C2[to_markdown]
        C3[to_markdown_with_layout]
        C4[to_json]
        C5[to_json_with_layout]
    end
    
    A --> B
    A1 --> B1
    A2 --> B2
    A3 --> B3
    
    B --> B1
    B --> B2
    B --> B3
    
    B3 --> B4
    B4 --> B5
    B4 --> B6
    B6 --> B7
    B7 --> B8
    
    B --> C
    
    C --> C1
    C --> C2
    C --> C3
    C --> C4
    C --> C5
    
    style A fill:#e1f5fe
    style C fill:#e8f5e9
    style B fill:#fff3e0
```

---

## 3. 데이터 구조 관계 (Data Structure Relationships)

```mermaid
flowchart LR
    subgraph Document["HwpxDocument"]
        direction TB
        DOC[📄 HwpxDocument]
        VER[VersionInfo]
        META[metadata]
    end
    
    subgraph Sections["Sections"]
        direction TB
        SEC[📑 Section]
        PAGE[PageProperties]
        MAR1[Margin]
    end
    
    subgraph Paragraphs["Paragraphs"]
        direction TB
        PARA[📝 Paragraph]
        LSEG[LineSegment]
        TRUN[TextRun]
    end
    
    subgraph Tables["Tables"]
        direction TB
        TBL[📊 Table]
        POS[Position]
        SIZE1[Size]
        MAR2[Margin]
    end
    
    subgraph Cells["Cells"]
        direction TB
        CELL[🔲 TableCell]
        SIZE2[Size]
        MAR3[Margin]
    end
    
    DOC --> VER
    DOC --> META
    DOC --> SEC
    
    SEC --> PAGE
    PAGE --> MAR1
    SEC --> PARA
    
    PARA --> LSEG
    PARA --> TRUN
    PARA --> TBL
    
    TBL --> POS
    TBL --> SIZE1
    TBL --> MAR2
    TBL --> CELL
    
    CELL --> SIZE2
    CELL --> MAR3
```

---

## 4. XML 파싱 상세 흐름 (XML Parsing Detail)

```mermaid
sequenceDiagram
    participant User as 사용자
    participant Parser as HwpxFolderParser
    participant ET as ElementTree
    participant Doc as HwpxDocument
    
    User->>Parser: parse_hwpx_folder(폴더경로)
    activate Parser
    
    Parser->>Parser: __init__(폴더경로)
    Note over Parser: 폴더 존재 확인<br/>Contents 폴더 확인
    
    Parser->>Parser: parse()
    
    rect rgb(255, 243, 224)
        Note over Parser,ET: 1단계: 버전 정보 파싱
        Parser->>ET: parse(version.xml)
        ET-->>Parser: root element
        Parser->>Doc: VersionInfo 생성
    end
    
    rect rgb(225, 245, 254)
        Note over Parser,ET: 2단계: 메타데이터 파싱
        Parser->>ET: parse(header.xml)
        ET-->>Parser: root element
        Parser->>Parser: 모든 요소 순회
        Parser->>Doc: metadata dict 생성
    end
    
    rect rgb(232, 245, 233)
        Note over Parser,ET: 3단계: 섹션 파싱
        Parser->>ET: parse(section0.xml)
        ET-->>Parser: root element
        
        Parser->>Parser: _parse_page_properties()
        Note over Parser: pagePr, margin 추출
        
        loop 각 p 요소
            Parser->>Parser: _parse_paragraph()
            Note over Parser: id, texts, line_segments 추출
            
            opt tbl 요소 발견
                Parser->>Parser: _parse_table()
                Note over Parser: sz, pos, margin 추출
                
                loop 각 tc 요소
                    Parser->>Parser: _parse_table_cell()
                    Note over Parser: cellSz, cellMargin 추출
                end
            end
        end
        
        Parser->>Doc: Section 추가
    end
    
    Parser-->>User: HwpxDocument 반환
    deactivate Parser
    
    User->>Doc: to_json_with_layout()
    Doc-->>User: JSON 문자열
    
    User->>Doc: to_markdown_with_layout()
    Doc-->>User: Markdown 문자열
```

---

## 5. 주요 클래스별 역할

```mermaid
mindmap
    root((HWPX Parser))
        레이아웃 클래스
            Position
                수평/수직 기준점
                정렬 방식
                오프셋 값
            Size
                너비/높이
                상대/절대 기준
            Margin
                상하좌우 여백
            LineSegment
                텍스트 라인 위치
                베이스라인
            PageProperties
                페이지 크기
                용지 방향
                페이지 여백
        콘텐츠 클래스
            HwpxDocument
                문서 전체
                출력 메서드들
            Section
                섹션 단위
                페이지 속성
            Paragraph
                문단 단위
                텍스트 런
                바운딩 박스
            Table
                테이블 구조
                셀 목록
            TableCell
                개별 셀
                병합 정보
            TextRun
                서식 단위
        파서 클래스
            HwpxFolderParser
                XML 파싱
                데이터 추출
                객체 생성
```

---

## 6. 출력 형식 비교

```mermaid
flowchart LR
    subgraph Input["입력"]
        DOC[HwpxDocument]
    end
    
    subgraph Basic["기본 출력"]
        T1[to_text]
        T2[to_markdown]
        T3[to_json]
    end
    
    subgraph Layout["레이아웃 포함 출력"]
        L1[to_markdown_with_layout]
        L2[to_json_with_layout]
    end
    
    subgraph Output["출력 내용"]
        O1["순수 텍스트만"]
        O2["텍스트 + 테이블 구조"]
        O3["텍스트 + 테이블<br/>(JSON 형식)"]
        O4["텍스트 + 테이블<br/>+ 좌표/크기 주석"]
        O5["모든 레이아웃 정보<br/>HWPUNIT + mm"]
    end
    
    DOC --> T1 --> O1
    DOC --> T2 --> O2
    DOC --> T3 --> O3
    DOC --> L1 --> O4
    DOC --> L2 --> O5
    
    style Basic fill:#e3f2fd
    style Layout fill:#fce4ec
```

---

## 7. 파일 구조와 클래스 매핑

```mermaid
flowchart TD
    subgraph HWPX["HWPX 폴더 구조"]
        F1["📁 hwpx_sample/"]
        F2["├── 📄 version.xml"]
        F3["├── 📄 settings.xml"]
        F4["├── 📁 Contents/"]
        F5["│   ├── 📄 header.xml"]
        F6["│   └── 📄 section0.xml"]
        F7["├── 📁 META-INF/"]
        F8["└── 📁 Preview/"]
    end
    
    subgraph Classes["파싱 결과 클래스"]
        C1["VersionInfo"]
        C2["(미사용)"]
        C3["metadata dict"]
        C4["Section"]
        C5["PageProperties"]
        C6["Paragraph"]
        C7["Table"]
        C8["TableCell"]
    end
    
    F2 --> C1
    F3 -.-> C2
    F5 --> C3
    F6 --> C4
    F6 --> C5
    F6 --> C6
    F6 --> C7
    F6 --> C8
    
    style HWPX fill:#fff8e1
    style Classes fill:#e8f5e9
```

---

## 8. 좌표 단위 변환

```mermaid
flowchart LR
    subgraph HWPUNIT["HWPUNIT (내부 단위)"]
        H1["width: 59528"]
        H2["height: 84186"]
        H3["margin: 2835"]
    end
    
    subgraph Conversion["변환 공식"]
        CONV["× 0.00353<br/>(25.4 / 7200)"]
    end
    
    subgraph MM["밀리미터 (mm)"]
        M1["width_mm: 210.0"]
        M2["height_mm: 296.99"]
        M3["margin_mm: 10.0"]
    end
    
    H1 --> CONV --> M1
    H2 --> CONV --> M2
    H3 --> CONV --> M3
    
    style HWPUNIT fill:#ffecb3
    style MM fill:#c8e6c9
```

---

## 사용 방법

### 다이어그램 보기

1. **GitHub**: 이 파일을 GitHub에 push하면 자동으로 렌더링됩니다.

2. **VS Code**: "Markdown Preview Mermaid Support" 확장 설치 후 미리보기

3. **온라인**: [Mermaid Live Editor](https://mermaid.live/)에 코드 복사

### 코드 사용 예시

```python
from hwpx_folder_parser_cursor import parse_hwpx_folder

# 1. 폴더 파싱
doc = parse_hwpx_folder("results/hwpx_sample")

# 2. 기본 텍스트 추출
text = doc.to_text()

# 3. 레이아웃 포함 JSON
json_data = doc.to_json_with_layout()

# 4. 레이아웃 포함 마크다운
markdown = doc.to_markdown_with_layout()
```



