"""
HWPX Folder Parser - 압축 해제된 HWPX 폴더의 XML 파싱 (레이아웃 정보 포함)

=============================================================================
HWPX 파일이란?
=============================================================================
HWPX는 한글(한컴오피스)에서 사용하는 문서 형식입니다.
실제로는 ZIP 파일이며, 압축을 풀면 여러 XML 파일들이 나옵니다.

HWPX 폴더 구조:
    hwpx_sample/
    ├── Contents/           # 실제 문서 내용이 들어있는 폴더
    │   ├── header.xml      # 문서 헤더 정보 (스타일, 폰트 등)
    │   ├── section0.xml    # 첫 번째 섹션의 본문 내용
    │   ├── section1.xml    # 두 번째 섹션 (있는 경우)
    │   └── content.hpf     # 콘텐츠 정보
    ├── META-INF/           # 메타 정보
    ├── Preview/            # 미리보기 이미지
    ├── settings.xml        # 설정 정보
    └── version.xml         # 버전 정보

좌표 단위:
    HWPX에서 사용하는 좌표 단위는 HWPUNIT입니다.
    1 HWPUNIT = 1/7200 인치 = 약 0.0035mm
    예: width="59528" → 약 210mm (A4 용지 너비)

이 파서는 압축이 이미 풀린 HWPX 폴더를 읽어서 텍스트와 레이아웃 정보를 추출합니다.
=============================================================================
"""

# =============================================================================
# 필요한 라이브러리 불러오기 (import)
# =============================================================================

# ┌─────────────────────────────────────────────────────────────────────────┐
# │              from __future__ import annotations 설명                     │
# └─────────────────────────────────────────────────────────────────────────┘
#
# 이 import는 Python의 타입 힌트(type hints)를 "미래 방식"으로 처리하게 합니다.
#
# ▶ 사용하는 이유:
#   1. Python 3.9 이전에서도 새로운 타입 힌트 문법을 사용할 수 있습니다.
#   2. 순환 참조 문제를 해결할 수 있습니다.
#
# ▶ 예시 1: 새로운 문법 사용
#
#   # 이 import 없이 Python 3.8에서:
#   def func(items: list[str]) -> dict[str, int]:  # ❌ 에러!
#       pass
#   
#   # 대신 이렇게 써야 함:
#   from typing import List, Dict
#   def func(items: List[str]) -> Dict[str, int]:  # ✅ 동작
#       pass
#   
#   # 이 import가 있으면 Python 3.8에서도:
#   from __future__ import annotations
#   def func(items: list[str]) -> dict[str, int]:  # ✅ 동작!
#       pass
#
# ▶ 예시 2: Union 타입
#
#   # 이 import 없이 Python 3.9 이전에서:
#   def func(path: str | Path):  # ❌ 에러!
#       pass
#   
#   # 이 import가 있으면:
#   from __future__ import annotations
#   def func(path: str | Path):  # ✅ 동작!
#       pass
#
# ▶ 작동 원리:
#   이 import가 있으면 모든 타입 힌트가 "문자열"로 저장됩니다.
#   실제로 타입을 평가하지 않으므로 아직 정의되지 않은 클래스도 참조 가능합니다.
#
# ▶ 주의사항:
#   - 반드시 파일의 첫 번째 import여야 합니다.
#   - Python 3.11+에서는 이것이 기본 동작이 될 예정입니다.
#
# =============================================================================

from __future__ import annotations  # 반드시 첫 번째 import!
import xml.etree.ElementTree as ET  # XML 파싱 라이브러리
from pathlib import Path            # 파일 경로 처리
from dataclasses import dataclass, field, asdict  # 데이터 클래스 관련
from typing import Iterator, Union, Optional, Any  # 타입 힌트
import json                         # JSON 변환


# =============================================================================
# XML 네임스페이스 정의
# =============================================================================
NS = {
    "sec": "http://www.hancom.co.kr/hwpml/2011/section",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hv": "http://www.hancom.co.kr/hwpml/2011/version",
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
}

# HWPUNIT을 mm로 변환하는 상수 (1 HWPUNIT = 1/7200 인치)
HWPUNIT_TO_MM = 25.4 / 7200  # 약 0.00353mm


# =============================================================================
# 레이아웃 관련 데이터 클래스
# =============================================================================
# 
# ┌─────────────────────────────────────────────────────────────────────────┐
# │                        @dataclass 데코레이터 설명                         │
# └─────────────────────────────────────────────────────────────────────────┘
#
# @dataclass는 Python 3.7+에서 제공하는 데코레이터로, "데이터를 담는 클래스"를
# 쉽게 만들 수 있게 해줍니다.
#
# ▶ @dataclass를 사용하는 이유:
#   1. __init__() 메서드를 자동으로 생성해줍니다.
#   2. __repr__() 메서드를 자동으로 생성해줍니다. (출력 시 보기 좋게)
#   3. __eq__() 메서드를 자동으로 생성해줍니다. (객체 비교 가능)
#   4. 코드가 간결해지고 읽기 쉬워집니다.
#
# ▶ @dataclass 없이 클래스를 만들면:
#
#   class Position:
#       def __init__(self, vert_rel_to="", horz_rel_to="", ...):
#           self.vert_rel_to = vert_rel_to
#           self.horz_rel_to = horz_rel_to
#           ...
#       
#       def __repr__(self):
#           return f"Position(vert_rel_to={self.vert_rel_to}, ...)"
#       
#       def __eq__(self, other):
#           return (self.vert_rel_to == other.vert_rel_to and ...)
#
# ▶ @dataclass를 사용하면:
#
#   @dataclass
#   class Position:
#       vert_rel_to: str = ""
#       horz_rel_to: str = ""
#       ...
#
#   → 위의 모든 메서드가 자동으로 생성됩니다!
#
# ▶ 사용 예시:
#   pos = Position(vert_rel_to="PARA", horz_rel_to="PARA")
#   print(pos)  # Position(vert_rel_to='PARA', horz_rel_to='PARA', ...)
#   pos2 = Position(vert_rel_to="PARA", horz_rel_to="PARA")
#   print(pos == pos2)  # True
#
# =============================================================================

@dataclass
class Position:
    """
    요소의 위치 정보를 담는 클래스
    
    HWPX에서 위치는 다양한 기준점(상대 위치)을 기준으로 지정됩니다.
    
    Attributes:
        vert_rel_to (str): 수직 위치 기준 (예: "PARA" = 문단 기준)
        horz_rel_to (str): 수평 위치 기준
        vert_align (str): 수직 정렬 (예: "TOP", "CENTER", "BOTTOM")
        horz_align (str): 수평 정렬 (예: "LEFT", "CENTER", "RIGHT")
        vert_offset (int): 수직 오프셋 (HWPUNIT)
        horz_offset (int): 수평 오프셋 (HWPUNIT)
        treat_as_char (bool): 글자처럼 취급할지 여부
        flow_with_text (bool): 텍스트와 함께 흐를지 여부
    """
    vert_rel_to: str = ""
    horz_rel_to: str = ""
    vert_align: str = ""
    horz_align: str = ""
    vert_offset: int = 0
    horz_offset: int = 0
    treat_as_char: bool = False
    flow_with_text: bool = False
    
    def to_mm(self) -> dict:
        """오프셋을 mm 단위로 변환하여 반환"""
        return {
            "vert_offset_mm": round(self.vert_offset * HWPUNIT_TO_MM, 2),
            "horz_offset_mm": round(self.horz_offset * HWPUNIT_TO_MM, 2),
        }


@dataclass
class Size:
    """
    요소의 크기 정보를 담는 클래스
    
    Attributes:
        width (int): 너비 (HWPUNIT)
        height (int): 높이 (HWPUNIT)
        width_rel_to (str): 너비 기준 (예: "ABSOLUTE", "PAPER", "PAGE")
        height_rel_to (str): 높이 기준
    """
    width: int = 0
    height: int = 0
    width_rel_to: str = "ABSOLUTE"
    height_rel_to: str = "ABSOLUTE"
    
    def to_mm(self) -> dict:
        """크기를 mm 단위로 변환하여 반환"""
        return {
            "width_mm": round(self.width * HWPUNIT_TO_MM, 2),
            "height_mm": round(self.height * HWPUNIT_TO_MM, 2),
        }


@dataclass
class Margin:
    """
    여백 정보를 담는 클래스
    
    Attributes:
        left (int): 왼쪽 여백 (HWPUNIT)
        right (int): 오른쪽 여백 (HWPUNIT)
        top (int): 위쪽 여백 (HWPUNIT)
        bottom (int): 아래쪽 여백 (HWPUNIT)
    """
    left: int = 0
    right: int = 0
    top: int = 0
    bottom: int = 0
    
    def to_mm(self) -> dict:
        """여백을 mm 단위로 변환하여 반환"""
        return {
            "left_mm": round(self.left * HWPUNIT_TO_MM, 2),
            "right_mm": round(self.right * HWPUNIT_TO_MM, 2),
            "top_mm": round(self.top * HWPUNIT_TO_MM, 2),
            "bottom_mm": round(self.bottom * HWPUNIT_TO_MM, 2),
        }


@dataclass 
class LineSegment:
    """
    텍스트 라인 세그먼트의 레이아웃 정보
    
    한 줄의 텍스트가 화면/종이에서 어디에 위치하는지를 나타냅니다.
    
    Attributes:
        text_pos (int): 텍스트 시작 위치 (문자 인덱스)
        vert_pos (int): 수직 위치 (HWPUNIT, 섹션 시작점 기준)
        vert_size (int): 수직 크기/줄 높이 (HWPUNIT)
        text_height (int): 실제 텍스트 높이 (HWPUNIT)
        baseline (int): 베이스라인 위치 (HWPUNIT)
        spacing (int): 줄 간격 (HWPUNIT)
        horz_pos (int): 수평 위치 (HWPUNIT)
        horz_size (int): 수평 크기/줄 너비 (HWPUNIT)
    """
    text_pos: int = 0
    vert_pos: int = 0
    vert_size: int = 0
    text_height: int = 0
    baseline: int = 0
    spacing: int = 0
    horz_pos: int = 0
    horz_size: int = 0
    
    def to_mm(self) -> dict:
        """좌표를 mm 단위로 변환하여 반환"""
        return {
            "vert_pos_mm": round(self.vert_pos * HWPUNIT_TO_MM, 2),
            "vert_size_mm": round(self.vert_size * HWPUNIT_TO_MM, 2),
            "horz_pos_mm": round(self.horz_pos * HWPUNIT_TO_MM, 2),
            "horz_size_mm": round(self.horz_size * HWPUNIT_TO_MM, 2),
        }


# ┌─────────────────────────────────────────────────────────────────────────┐
# │                    field(default_factory=...) 설명                       │
# └─────────────────────────────────────────────────────────────────────────┘
#
# field()는 dataclass의 필드(속성)에 대한 추가 설정을 할 때 사용합니다.
# 특히 default_factory는 "가변 객체"를 기본값으로 사용할 때 필수입니다.
#
# ▶ 문제 상황 (default_factory를 안 쓰면):
#
#   @dataclass
#   class PageProperties:
#       margin: Margin = Margin()  # ❌ 위험! 모든 인스턴스가 같은 객체를 공유!
#
#   page1 = PageProperties()
#   page2 = PageProperties()
#   page1.margin.left = 100
#   print(page2.margin.left)  # 100 출력! (예상: 0)
#   # page1과 page2가 같은 Margin 객체를 공유하기 때문!
#
# ▶ 해결 (default_factory 사용):
#
#   @dataclass
#   class PageProperties:
#       margin: Margin = field(default_factory=Margin)  # ✅ 안전!
#
#   page1 = PageProperties()
#   page2 = PageProperties()
#   page1.margin.left = 100
#   print(page2.margin.left)  # 0 출력 (정상)
#   # 각 인스턴스마다 새로운 Margin 객체가 생성됨!
#
# ▶ default_factory에는 "호출 가능한 객체"를 전달합니다:
#   - field(default_factory=list)   → 빈 리스트 []
#   - field(default_factory=dict)   → 빈 딕셔너리 {}
#   - field(default_factory=Margin) → 새 Margin 객체
#   - field(default_factory=lambda: [1, 2, 3]) → [1, 2, 3] 리스트
#
# ▶ 언제 사용해야 하나요?
#   - 기본값이 리스트, 딕셔너리, 클래스 인스턴스 등 "가변 객체"일 때
#   - 기본값이 정수, 문자열, 불리언 등 "불변 객체"일 때는 필요 없음
#
# =============================================================================

@dataclass
class PageProperties:
    """
    페이지 속성 정보
    
    Attributes:
        width (int): 페이지 너비 (HWPUNIT)
        height (int): 페이지 높이 (HWPUNIT)
        landscape (str): 용지 방향 ("WIDELY" = 가로, "NARROWLY" = 세로)
        gutter_type (str): 제본 여백 위치
        margin (Margin): 페이지 여백
    """
    width: int = 0                                    # 불변 객체 → 그냥 기본값 사용 OK
    height: int = 0                                   # 불변 객체 → 그냥 기본값 사용 OK
    landscape: str = "NARROWLY"                       # 불변 객체 → 그냥 기본값 사용 OK
    gutter_type: str = "LEFT_ONLY"                    # 불변 객체 → 그냥 기본값 사용 OK
    margin: Margin = field(default_factory=Margin)    # 가변 객체 → field() 필수!
    
    def to_mm(self) -> dict:
        """페이지 크기를 mm 단위로 변환"""
        return {
            "width_mm": round(self.width * HWPUNIT_TO_MM, 2),
            "height_mm": round(self.height * HWPUNIT_TO_MM, 2),
            "orientation": "가로" if self.landscape == "WIDELY" else "세로",
        }


# =============================================================================
# 콘텐츠 데이터 클래스
# =============================================================================

@dataclass
class TableCell:
    """
    테이블의 한 칸(셀)을 나타내는 클래스
    
    Attributes:
        row (int): 행 번호 (0부터 시작)
        col (int): 열 번호 (0부터 시작)
        text (str): 셀 안에 들어있는 텍스트
        row_span (int): 세로로 합쳐진 셀 개수
        col_span (int): 가로로 합쳐진 셀 개수
        size (Size): 셀 크기
        margin (Margin): 셀 내부 여백
        border_fill_id (str): 테두리/배경 스타일 참조 ID
    """
    row: int
    col: int
    text: str
    row_span: int = 1
    col_span: int = 1
    size: Size = field(default_factory=Size)
    margin: Margin = field(default_factory=Margin)
    border_fill_id: str = ""


@dataclass
class Table:
    """
    테이블 전체를 나타내는 클래스
    
    Attributes:
        rows (int): 테이블의 총 행 수
        cols (int): 테이블의 총 열 수
        cells (list[TableCell]): 테이블에 포함된 모든 셀 목록
        id (str): 테이블 고유 ID
        z_order (int): 겹침 순서 (높을수록 위에 표시)
        position (Position): 테이블 위치
        size (Size): 테이블 크기
        outer_margin (Margin): 테이블 외부 여백
        inner_margin (Margin): 테이블 내부 여백 (셀과 테두리 사이)
    """
    rows: int
    cols: int
    cells: list[TableCell] = field(default_factory=list)
    id: str = ""
    z_order: int = 0
    position: Position = field(default_factory=Position)
    size: Size = field(default_factory=Size)
    outer_margin: Margin = field(default_factory=Margin)
    inner_margin: Margin = field(default_factory=Margin)
    
    def to_markdown(self) -> str:
        """마크다운 테이블로 변환 (레이아웃 정보 제외)"""
        if not self.cells:
            return ""
        
        grid = [["" for _ in range(self.cols)] for _ in range(self.rows)]
        for cell in self.cells:
            if 0 <= cell.row < self.rows and 0 <= cell.col < self.cols:
                grid[cell.row][cell.col] = cell.text.replace("|", "\\|")
        
        lines = []
        for i, row in enumerate(grid):
            lines.append("| " + " | ".join(row) + " |")
            if i == 0:
                lines.append("|" + "|".join(["---"] * self.cols) + "|")
        
        return "\n".join(lines)
    
    def to_markdown_with_layout(self) -> str:
        """레이아웃 정보를 포함한 마크다운 테이블"""
        lines = []
        
        # 테이블 메타 정보
        size_mm = self.size.to_mm()
        lines.append(f"<!-- 테이블 ID: {self.id} -->")
        lines.append(f"<!-- 크기: {size_mm['width_mm']}mm × {size_mm['height_mm']}mm -->")
        lines.append(f"<!-- 위치: {self.position.horz_align} / {self.position.vert_align} -->")
        lines.append("")
        
        # 테이블 본문
        lines.append(self.to_markdown())
        
        return "\n".join(lines)


@dataclass
class TextRun:
    """
    텍스트 런(run) - 동일한 서식이 적용된 텍스트 조각
    
    Attributes:
        text (str): 텍스트 내용
        char_pr_id (str): 문자 속성 참조 ID (폰트, 크기 등)
    """
    text: str
    char_pr_id: str = ""


@dataclass
class Paragraph:
    """
    문단(Paragraph)을 나타내는 클래스
    
    Attributes:
        id (str): 문단의 고유 식별자
        texts (list[str]): 문단에 포함된 텍스트 조각들
        text_runs (list[TextRun]): 서식 정보가 포함된 텍스트 런들
        tables (list[Table]): 문단에 포함된 테이블들
        para_pr_id (str): 문단 속성 참조 ID
        style_id (str): 스타일 참조 ID
        line_segments (list[LineSegment]): 라인 세그먼트 레이아웃 정보
        page_break (bool): 문단 앞에서 페이지 나누기
        column_break (bool): 문단 앞에서 단 나누기
    """
    id: str
    texts: list[str] = field(default_factory=list)
    text_runs: list[TextRun] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    para_pr_id: str = ""
    style_id: str = ""
    line_segments: list[LineSegment] = field(default_factory=list)
    page_break: bool = False
    column_break: bool = False
    
    # ─────────────────────────────────────────────────────────────────────────
    # @property 데코레이터 설명
    # ─────────────────────────────────────────────────────────────────────────
    #
    # @property는 메서드를 "속성처럼" 사용할 수 있게 해주는 데코레이터입니다.
    #
    # ▶ @property를 사용하는 이유:
    #   1. 메서드 호출 시 괄호()를 생략할 수 있어 코드가 깔끔해집니다.
    #   2. 계산된 값을 속성처럼 접근할 수 있습니다.
    #   3. 내부 구현을 숨기면서 간단한 인터페이스를 제공합니다.
    #
    # ▶ @property 없이 사용하면:
    #   text = para.full_text()  # 괄호 필요
    #
    # ▶ @property를 사용하면:
    #   text = para.full_text    # 괄호 없이 속성처럼 접근!
    #
    # ▶ 언제 사용하나요?
    #   - 값을 "읽기 전용"으로 제공하고 싶을 때
    #   - 계산된 값을 속성처럼 접근하게 하고 싶을 때
    #   - 예: full_text, title, bounding_box 등
    #
    # ▶ 주의사항:
    #   - @property는 읽기 전용입니다. 값을 변경하려면 @setter도 정의해야 합니다.
    #   - 무거운 계산은 @property보다 일반 메서드가 적합합니다.
    #
    # ─────────────────────────────────────────────────────────────────────────
    
    @property
    def full_text(self) -> str:
        """
        문단의 전체 텍스트를 하나의 문자열로 반환
        
        사용 예시:
            para = Paragraph(id="1", texts=["안녕", "하세요"])
            print(para.full_text)  # "안녕하세요" (괄호 없이 속성처럼 접근!)
        """
        return "".join(self.texts)
    
    def get_bounding_box(self) -> dict | None:
        """문단의 바운딩 박스(위치와 크기)를 반환"""
        if not self.line_segments:
            return None
        
        # 모든 라인 세그먼트의 범위 계산
        min_x = min(ls.horz_pos for ls in self.line_segments)
        max_x = max(ls.horz_pos + ls.horz_size for ls in self.line_segments)
        min_y = min(ls.vert_pos for ls in self.line_segments)
        max_y = max(ls.vert_pos + ls.vert_size for ls in self.line_segments)
        
        return {
            "x": min_x,
            "y": min_y,
            "width": max_x - min_x,
            "height": max_y - min_y,
            "x_mm": round(min_x * HWPUNIT_TO_MM, 2),
            "y_mm": round(min_y * HWPUNIT_TO_MM, 2),
            "width_mm": round((max_x - min_x) * HWPUNIT_TO_MM, 2),
            "height_mm": round((max_y - min_y) * HWPUNIT_TO_MM, 2),
        }


@dataclass
class Section:
    """
    섹션(Section)을 나타내는 클래스
    
    Attributes:
        index (int): 섹션 번호 (0부터 시작)
        paragraphs (list[Paragraph]): 섹션에 포함된 문단들
        page_props (PageProperties): 페이지 속성 (크기, 여백 등)
    """
    index: int
    paragraphs: list[Paragraph] = field(default_factory=list)
    page_props: PageProperties = field(default_factory=PageProperties)
    
    @property
    def full_text(self) -> str:
        """섹션의 전체 텍스트를 반환"""
        return "\n".join(p.full_text for p in self.paragraphs if p.full_text)


@dataclass
class VersionInfo:
    """HWPX 파일의 버전 정보"""
    application: str = ""
    app_version: str = ""
    xml_version: str = ""


@dataclass
class HwpxDocument:
    """
    HWPX 문서 전체를 나타내는 클래스
    
    Attributes:
        folder_path (Path): HWPX 폴더 경로
        version (VersionInfo): 버전 정보
        sections (list[Section]): 문서의 모든 섹션
        metadata (dict): 문서 메타데이터
    """
    folder_path: Path
    version: VersionInfo = field(default_factory=VersionInfo)
    sections: list[Section] = field(default_factory=list)
    metadata: dict = field(default_factory=dict)
    
    @property
    def title(self) -> str:
        """문서 제목 (폴더명)"""
        return self.folder_path.name
    
    def to_text(self) -> str:
        """문서 전체의 텍스트만 추출 (레이아웃 정보 제외)"""
        return "\n\n".join(s.full_text for s in self.sections if s.full_text)
    
    def to_markdown(self) -> str:
        """문서를 마크다운 형식으로 변환 (기본 버전)"""
        lines = [f"# {self.title}", ""]
        
        for section in self.sections:
            lines.append(f"## Section {section.index + 1}")
            lines.append("")
            
            for para in section.paragraphs:
                if para.full_text:
                    lines.append(para.full_text)
                    lines.append("")
                
                for table in para.tables:
                    lines.append(table.to_markdown())
                    lines.append("")
        
        return "\n".join(lines)
    
    def to_markdown_with_layout(self) -> str:
        """
        레이아웃 정보를 포함한 마크다운
        
        HTML 주석으로 좌표/크기 정보를 포함합니다.
        """
        lines = [f"# {self.title}", ""]
        
        # 문서 메타 정보
        lines.append("<!-- 문서 정보 -->")
        lines.append(f"<!-- 버전: {self.version.application} {self.version.app_version} -->")
        lines.append("")
        
        for section in self.sections:
            # 섹션 헤더
            lines.append(f"## Section {section.index + 1}")
            lines.append("")
            
            # 페이지 정보
            page_mm = section.page_props.to_mm()
            lines.append(f"<!-- 페이지 크기: {page_mm['width_mm']}mm × {page_mm['height_mm']}mm ({page_mm['orientation']}) -->")
            margin_mm = section.page_props.margin.to_mm()
            lines.append(f"<!-- 여백: 상{margin_mm['top_mm']}mm 하{margin_mm['bottom_mm']}mm 좌{margin_mm['left_mm']}mm 우{margin_mm['right_mm']}mm -->")
            lines.append("")
            
            for para in section.paragraphs:
                # 문단 레이아웃 정보
                bbox = para.get_bounding_box()
                if bbox:
                    lines.append(f"<!-- 문단 위치: ({bbox['x_mm']}mm, {bbox['y_mm']}mm), 크기: {bbox['width_mm']}mm × {bbox['height_mm']}mm -->")
                
                if para.full_text:
                    lines.append(para.full_text)
                    lines.append("")
                
                for table in para.tables:
                    lines.append(table.to_markdown_with_layout())
                    lines.append("")
        
        return "\n".join(lines)
    
    def to_json(self) -> str:
        """문서를 JSON 형식으로 변환 (기본 버전, 레이아웃 정보 제외)"""
        # 기본 정보만 포함
        data = {
            "title": self.title,
            "version": asdict(self.version),
            "sections": [
                {
                    "index": s.index,
                    "paragraphs": [
                        {
                            "id": p.id,
                            "text": p.full_text,
                            "tables": [
                                {
                                    "rows": t.rows,
                                    "cols": t.cols,
                                    "cells": [
                                        {"row": c.row, "col": c.col, "text": c.text}
                                        for c in t.cells
                                    ]
                                }
                                for t in p.tables
                            ]
                        }
                        for p in s.paragraphs
                        if p.full_text or p.tables
                    ]
                }
                for s in self.sections
            ]
        }
        return json.dumps(data, ensure_ascii=False, indent=2)
    
    def to_json_with_layout(self) -> str:
        """
        레이아웃 정보를 포함한 JSON
        
        좌표, 크기, 여백 등의 정보가 모두 포함됩니다.
        HWPUNIT과 mm 단위 모두 제공합니다.
        """
        data = {
            "title": self.title,
            "version": asdict(self.version),
            "metadata": self.metadata,
            "unit_info": {
                "description": "좌표 단위 정보",
                "hwpunit": "1 HWPUNIT = 1/7200 인치 ≈ 0.00353mm",
                "conversion": HWPUNIT_TO_MM,
            },
            "sections": []
        }
        
        for section in self.sections:
            section_data = {
                "index": section.index,
                "page_properties": {
                    **asdict(section.page_props),
                    "size_mm": section.page_props.to_mm(),
                    "margin_mm": section.page_props.margin.to_mm(),
                },
                "paragraphs": []
            }
            
            for para in section.paragraphs:
                if not para.full_text and not para.tables:
                    continue
                    
                para_data = {
                    "id": para.id,
                    "text": para.full_text,
                    "style": {
                        "para_pr_id": para.para_pr_id,
                        "style_id": para.style_id,
                    },
                    "layout": {
                        "page_break": para.page_break,
                        "column_break": para.column_break,
                        "bounding_box": para.get_bounding_box(),
                        "line_segments": [
                            {
                                **asdict(ls),
                                "position_mm": ls.to_mm(),
                            }
                            for ls in para.line_segments
                        ]
                    },
                    "text_runs": [
                        {"text": tr.text, "char_pr_id": tr.char_pr_id}
                        for tr in para.text_runs
                    ],
                    "tables": []
                }
                
                for table in para.tables:
                    table_data = {
                        "id": table.id,
                        "rows": table.rows,
                        "cols": table.cols,
                        "z_order": table.z_order,
                        "layout": {
                            "position": {
                                **asdict(table.position),
                                "offset_mm": table.position.to_mm(),
                            },
                            "size": {
                                **asdict(table.size),
                                "size_mm": table.size.to_mm(),
                            },
                            "outer_margin": {
                                **asdict(table.outer_margin),
                                "margin_mm": table.outer_margin.to_mm(),
                            },
                            "inner_margin": {
                                **asdict(table.inner_margin),
                                "margin_mm": table.inner_margin.to_mm(),
                            },
                        },
                        "cells": [
                            {
                                "row": c.row,
                                "col": c.col,
                                "text": c.text,
                                "row_span": c.row_span,
                                "col_span": c.col_span,
                                "border_fill_id": c.border_fill_id,
                                "size": {
                                    **asdict(c.size),
                                    "size_mm": c.size.to_mm(),
                                },
                                "margin": {
                                    **asdict(c.margin),
                                    "margin_mm": c.margin.to_mm(),
                                },
                            }
                            for c in table.cells
                        ]
                    }
                    para_data["tables"].append(table_data)
                
                section_data["paragraphs"].append(para_data)
            
            data["sections"].append(section_data)
        
        return json.dumps(data, ensure_ascii=False, indent=2)


# =============================================================================
# HWPX 폴더 파서 클래스
# =============================================================================

class HwpxFolderParser:
    """
    압축 해제된 HWPX 폴더를 파싱하는 클래스
    
    사용법:
        parser = HwpxFolderParser("results/hwpx_sample")
        doc = parser.parse()
        
        # 기본 텍스트
        print(doc.to_text())
        
        # 레이아웃 포함 JSON
        print(doc.to_json_with_layout())
        
        # 레이아웃 포함 마크다운
        print(doc.to_markdown_with_layout())
    """
    
    def __init__(self, folder_path: str | Path):
        """
        파서를 초기화합니다.
        
        Args:
            folder_path: HWPX 폴더 경로
        
        Raises:
            FileNotFoundError: 폴더가 존재하지 않거나 Contents 폴더가 없을 때
        """
        self.folder_path = Path(folder_path)
        
        if not self.folder_path.exists():
            raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {folder_path}")
        
        self.contents_dir = self.folder_path / "Contents"
        if not self.contents_dir.exists():
            raise FileNotFoundError(f"Contents 폴더를 찾을 수 없습니다: {self.contents_dir}")
    
    def parse(self) -> HwpxDocument:
        """HWPX 폴더 전체를 파싱합니다."""
        doc = HwpxDocument(folder_path=self.folder_path)
        
        doc.version = self._parse_version()
        doc.metadata = self._parse_metadata()
        doc.sections = list(self._parse_sections())
        
        return doc
    
    def _parse_version(self) -> VersionInfo:
        """version.xml 파일을 파싱"""
        version_file = self.folder_path / "version.xml"
        info = VersionInfo()
        
        if version_file.exists():
            try:
                tree = ET.parse(version_file)
                root = tree.getroot()
                info.application = root.get("application", "")
                info.app_version = root.get("appVersion", "")
                info.xml_version = root.get("xmlVersion", "")
            except ET.ParseError:
                pass
        
        return info
    
    def _parse_metadata(self) -> dict:
        """header.xml에서 메타데이터 추출"""
        header_file = self.contents_dir / "header.xml"
        meta = {}
        
        if header_file.exists():
            try:
                tree = ET.parse(header_file)
                root = tree.getroot()
                
                for elem in root.iter():
                    tag = self._strip_ns(elem.tag)
                    
                    if elem.attrib:
                        if tag not in meta:
                            meta[tag] = []
                        meta[tag].append(dict(elem.attrib))
                    
                    if elem.text and elem.text.strip():
                        meta[f"{tag}_text"] = elem.text.strip()
                        
            except ET.ParseError:
                pass
        
        return meta
    
    def _parse_sections(self) -> Iterator[Section]:
        """모든 섹션 파싱"""
        section_files = sorted(self.contents_dir.glob("section*.xml"))
        
        for idx, section_file in enumerate(section_files):
            yield self._parse_section(section_file, idx)
    
    def _parse_section(self, section_file: Path, index: int) -> Section:
        """단일 섹션 파싱"""
        section = Section(index=index)
        
        try:
            tree = ET.parse(section_file)
            root = tree.getroot()
            
            # 페이지 속성 추출
            section.page_props = self._parse_page_properties(root)
            
            # 문단 추출
            for p_elem in root.iter():
                if self._strip_ns(p_elem.tag) == "p":
                    para = self._parse_paragraph(p_elem)
                    if para.texts or para.tables:
                        section.paragraphs.append(para)
                        
        except ET.ParseError as e:
            print(f"XML 파싱 오류 ({section_file}): {e}")
        
        return section
    
    def _parse_page_properties(self, root) -> PageProperties:
        """페이지 속성 추출"""
        props = PageProperties()
        
        # pagePr 요소 찾기
        for elem in root.iter():
            tag = self._strip_ns(elem.tag)
            
            if tag == "pagePr":
                props.width = int(elem.get("width", 0))
                props.height = int(elem.get("height", 0))
                props.landscape = elem.get("landscape", "NARROWLY")
                props.gutter_type = elem.get("gutterType", "LEFT_ONLY")
            
            elif tag == "margin":
                props.margin = Margin(
                    left=int(elem.get("left", 0)),
                    right=int(elem.get("right", 0)),
                    top=int(elem.get("top", 0)),
                    bottom=int(elem.get("bottom", 0))
                )
        
        return props
    
    def _parse_paragraph(self, p_elem) -> Paragraph:
        """문단 파싱 (레이아웃 정보 포함)"""
        para = Paragraph(
            id=p_elem.get("id", ""),
            para_pr_id=p_elem.get("paraPrIDRef", ""),
            style_id=p_elem.get("styleIDRef", ""),
            page_break=p_elem.get("pageBreak", "0") == "1",
            column_break=p_elem.get("columnBreak", "0") == "1",
        )
        
        for elem in p_elem.iter():
            tag = self._strip_ns(elem.tag)
            
            # 텍스트 런 추출
            if tag == "run":
                char_pr_id = elem.get("charPrIDRef", "")
                for child in elem.iter():
                    if self._strip_ns(child.tag) == "t" and child.text:
                        para.texts.append(child.text)
                        para.text_runs.append(TextRun(text=child.text, char_pr_id=char_pr_id))
            
            # 라인 세그먼트 추출
            elif tag == "lineseg":
                ls = LineSegment(
                    text_pos=int(elem.get("textpos", 0)),
                    vert_pos=int(elem.get("vertpos", 0)),
                    vert_size=int(elem.get("vertsize", 0)),
                    text_height=int(elem.get("textheight", 0)),
                    baseline=int(elem.get("baseline", 0)),
                    spacing=int(elem.get("spacing", 0)),
                    horz_pos=int(elem.get("horzpos", 0)),
                    horz_size=int(elem.get("horzsize", 0)),
                )
                para.line_segments.append(ls)
            
            # 테이블 추출
            elif tag == "tbl":
                table = self._parse_table(elem)
                if table:
                    para.tables.append(table)
        
        return para
    
    def _parse_table(self, tbl_elem) -> Table | None:
        """테이블 파싱 (레이아웃 정보 포함)"""
        rows = int(tbl_elem.get("rowCnt", 0))
        cols = int(tbl_elem.get("colCnt", 0))
        
        if rows == 0 or cols == 0:
            return None
        
        table = Table(
            rows=rows,
            cols=cols,
            id=tbl_elem.get("id", ""),
            z_order=int(tbl_elem.get("zOrder", 0)),
        )
        
        # 테이블 레이아웃 정보 추출
        for elem in tbl_elem:
            tag = self._strip_ns(elem.tag)
            
            if tag == "sz":
                table.size = Size(
                    width=int(elem.get("width", 0)),
                    height=int(elem.get("height", 0)),
                    width_rel_to=elem.get("widthRelTo", "ABSOLUTE"),
                    height_rel_to=elem.get("heightRelTo", "ABSOLUTE"),
                )
            
            elif tag == "pos":
                table.position = Position(
                    vert_rel_to=elem.get("vertRelTo", ""),
                    horz_rel_to=elem.get("horzRelTo", ""),
                    vert_align=elem.get("vertAlign", ""),
                    horz_align=elem.get("horzAlign", ""),
                    vert_offset=int(elem.get("vertOffset", 0)),
                    horz_offset=int(elem.get("horzOffset", 0)),
                    treat_as_char=elem.get("treatAsChar", "0") == "1",
                    flow_with_text=elem.get("flowWithText", "0") == "1",
                )
            
            elif tag == "outMargin":
                table.outer_margin = Margin(
                    left=int(elem.get("left", 0)),
                    right=int(elem.get("right", 0)),
                    top=int(elem.get("top", 0)),
                    bottom=int(elem.get("bottom", 0)),
                )
            
            elif tag == "inMargin":
                table.inner_margin = Margin(
                    left=int(elem.get("left", 0)),
                    right=int(elem.get("right", 0)),
                    top=int(elem.get("top", 0)),
                    bottom=int(elem.get("bottom", 0)),
                )
        
        # 셀 추출
        row_idx = 0
        for elem in tbl_elem.iter():
            tag = self._strip_ns(elem.tag)
            
            if tag == "tr":
                col_idx = 0
                for cell_elem in elem.iter():
                    if self._strip_ns(cell_elem.tag) == "tc":
                        cell = self._parse_table_cell(cell_elem, row_idx, col_idx)
                        table.cells.append(cell)
                        col_idx += 1
                row_idx += 1
        
        return table
    
    def _parse_table_cell(self, tc_elem, row: int, col: int) -> TableCell:
        """테이블 셀 파싱"""
        # 셀 텍스트 추출
        cell_texts = []
        for t_elem in tc_elem.iter():
            if self._strip_ns(t_elem.tag) == "t" and t_elem.text:
                cell_texts.append(t_elem.text)
        
        cell = TableCell(
            row=row,
            col=col,
            text=" ".join(cell_texts),
            border_fill_id=tc_elem.get("borderFillIDRef", ""),
        )
        
        # 셀 레이아웃 정보 추출
        for elem in tc_elem.iter():
            tag = self._strip_ns(elem.tag)
            
            if tag == "cellSpan":
                cell.row_span = int(elem.get("rowSpan", 1))
                cell.col_span = int(elem.get("colSpan", 1))
            
            elif tag == "cellSz":
                cell.size = Size(
                    width=int(elem.get("width", 0)),
                    height=int(elem.get("height", 0)),
                )
            
            elif tag == "cellMargin":
                cell.margin = Margin(
                    left=int(elem.get("left", 0)),
                    right=int(elem.get("right", 0)),
                    top=int(elem.get("top", 0)),
                    bottom=int(elem.get("bottom", 0)),
                )
        
        return cell
    
    # ─────────────────────────────────────────────────────────────────────────
    # @staticmethod 데코레이터 설명
    # ─────────────────────────────────────────────────────────────────────────
    #
    # @staticmethod는 클래스의 인스턴스(self) 없이 호출할 수 있는 메서드를 정의합니다.
    #
    # ▶ @staticmethod를 사용하는 이유:
    #   1. 메서드가 인스턴스 변수(self.xxx)를 전혀 사용하지 않을 때
    #   2. 클래스와 관련있지만 독립적인 유틸리티 함수일 때
    #   3. 코드의 의도를 명확하게 표현할 수 있음
    #
    # ▶ 일반 메서드 vs 정적 메서드:
    #
    #   # 일반 메서드 - self 필요
    #   def _strip_ns(self, tag: str) -> str:
    #       return tag.split("}")[-1]
    #   
    #   parser = HwpxFolderParser("folder")
    #   result = parser._strip_ns("{ns}tag")  # 인스턴스 필요
    #
    #   # 정적 메서드 - self 불필요
    #   @staticmethod
    #   def _strip_ns(tag: str) -> str:
    #       return tag.split("}")[-1]
    #   
    #   result = HwpxFolderParser._strip_ns("{ns}tag")  # 인스턴스 없이 호출 가능!
    #   # 또는
    #   parser = HwpxFolderParser("folder")
    #   result = parser._strip_ns("{ns}tag")  # 인스턴스로도 호출 가능
    #
    # ▶ 언제 사용하나요?
    #   - self를 전혀 사용하지 않는 메서드일 때
    #   - 순수 함수(입력만으로 출력이 결정되는 함수)일 때
    #   - 클래스 외부에 둘 수도 있지만, 클래스와 관련있어 묶어두고 싶을 때
    #
    # ▶ @staticmethod vs @classmethod:
    #   - @staticmethod: 클래스/인스턴스 정보가 전혀 필요 없을 때
    #   - @classmethod: 클래스 자체(cls)에 접근해야 할 때 (예: 팩토리 메서드)
    #
    # ─────────────────────────────────────────────────────────────────────────
    
    @staticmethod
    def _strip_ns(tag: str) -> str:
        """
        XML 태그에서 네임스페이스를 제거합니다.
        
        예시:
            _strip_ns("{http://www.hancom.co.kr/hwpml}tag")
            → "tag"
        
        이 메서드는 self를 사용하지 않으므로 @staticmethod로 정의합니다.
        클래스 이름으로 직접 호출 가능: HwpxFolderParser._strip_ns(tag)
        """
        if "}" in tag:
            return tag.split("}")[-1]
        return tag


# =============================================================================
# 편의 함수
# =============================================================================

def parse_hwpx_folder(folder_path: str | Path) -> HwpxDocument:
    """HWPX 폴더를 파싱하는 편의 함수"""
    parser = HwpxFolderParser(folder_path)
    return parser.parse()


def parse_hwpx(file_path: str | Path, extract_dir: str | Path | None = None) -> HwpxDocument:
    """
    HWPX 파일(.hwpx)을 직접 파싱하는 함수
    
    HWPX 파일은 ZIP 형식이므로 먼저 압축을 해제하고 파싱합니다.
    
    Args:
        file_path: HWPX 파일 경로
        extract_dir: 압축 해제할 디렉토리 (None이면 임시 디렉토리 사용)
    
    Returns:
        HwpxDocument: 파싱된 문서 객체
    
    사용 예시:
        doc = parse_hwpx("document.hwpx")
        print(doc.to_text())
    """
    import zipfile
    import tempfile
    import shutil
    
    file_path = Path(file_path)
    
    if not file_path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
    
    if not file_path.suffix.lower() == ".hwpx":
        raise ValueError(f"HWPX 파일이 아닙니다: {file_path}")
    
    # 압축 해제 디렉토리 결정
    if extract_dir is None:
        # 임시 디렉토리 사용
        temp_dir = tempfile.mkdtemp(prefix="hwpx_")
        extract_path = Path(temp_dir) / file_path.stem
        cleanup = True
    else:
        extract_path = Path(extract_dir) / file_path.stem
        cleanup = False
    
    try:
        # ZIP 압축 해제
        with zipfile.ZipFile(file_path, 'r') as zf:
            zf.extractall(extract_path)
        
        # 폴더 파싱
        doc = parse_hwpx_folder(extract_path)
        
        # 원본 파일 경로 저장 (폴더 경로 대신)
        doc.folder_path = file_path
        
        return doc
        
    finally:
        # 임시 디렉토리 정리 (필요한 경우)
        if cleanup and extract_path.exists():
            shutil.rmtree(extract_path.parent, ignore_errors=True)


# =============================================================================
# 메인 실행부
# =============================================================================

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        folder = sys.argv[1]
    else:
        folder = "results/hwpx_sample"
    
    print(f"파싱 중: {folder}")
    print("=" * 60)
    
    doc = parse_hwpx_folder(folder)
    
    # 기본 정보 출력
    print(f"\n📄 문서: {doc.title}")
    print(f"📋 버전: {doc.version.application} {doc.version.app_version}")
    print(f"📑 섹션 수: {len(doc.sections)}")
    
    for section in doc.sections:
        print(f"\n--- Section {section.index + 1} ---")
        print(f"  문단 수: {len(section.paragraphs)}")
        
        # 페이지 정보
        page_mm = section.page_props.to_mm()
        print(f"  페이지: {page_mm['width_mm']}mm × {page_mm['height_mm']}mm ({page_mm['orientation']})")
        
        table_count = sum(len(p.tables) for p in section.paragraphs)
        if table_count:
            print(f"  테이블 수: {table_count}")
    
    # 레이아웃 포함 마크다운 출력
    print("\n" + "=" * 60)
    print("📐 레이아웃 포함 마크다운 (처음 2000자):")
    print("=" * 60)
    md = doc.to_markdown_with_layout()
    print(md[:2000] if len(md) > 2000 else md)
    
    # 레이아웃 포함 JSON 일부 출력
    print("\n" + "=" * 60)
    print("📊 레이아웃 포함 JSON (처음 3000자):")
    print("=" * 60)
    json_str = doc.to_json_with_layout()
    print(json_str[:3000] if len(json_str) > 3000 else json_str)
    
    # 파일로 저장
    output_dir = Path(folder).parent
    
    # JSON 저장
    json_file = output_dir / f"{doc.title}_layout.json"
    with open(json_file, "w", encoding="utf-8") as f:
        f.write(doc.to_json_with_layout())
    print(f"\n✅ JSON 저장: {json_file}")
    
    # 마크다운 저장
    md_file = output_dir / f"{doc.title}_layout.md"
    with open(md_file, "w", encoding="utf-8") as f:
        f.write(doc.to_markdown_with_layout())
    print(f"✅ 마크다운 저장: {md_file}")


# =============================================================================
# 레이아웃 정보 추출 및 시각화 함수
# =============================================================================
#
# 아래 함수들은 문서에서 레이아웃 정보를 추출하고 시각화하는 기능을 제공합니다.
#
# 1. extract_layout_elements(): 문서에서 필수 레이아웃 정보 추출
# 2. visualize_document(): 문서를 화이트보드에 시각화
# 3. create_document_viewer(): 인터랙티브 뷰어 생성
#
# =============================================================================

@dataclass
class LayoutElement:
    """
    레이아웃 요소 - 화면에 그릴 수 있는 단위
    
    Attributes:
        element_type (str): 요소 유형 ("text", "table", "table_cell")
        text (str): 텍스트 내용
        x (float): X 좌표 (mm)
        y (float): Y 좌표 (mm)
        width (float): 너비 (mm)
        height (float): 높이 (mm)
        page (int): 페이지 번호 (0부터 시작)
        section (int): 섹션 번호
        para_id (str): 문단 ID
        style_id (str): 스타일 ID
        metadata (dict): 추가 메타데이터
    """
    element_type: str
    text: str
    x: float
    y: float
    width: float
    height: float
    page: int = 0
    section: int = 0
    para_id: str = ""
    style_id: str = ""
    metadata: dict = field(default_factory=dict)
    
    def to_dict(self) -> dict:
        """딕셔너리로 변환"""
        return {
            "type": self.element_type,
            "text": self.text,
            "bbox": {
                "x": self.x,
                "y": self.y,
                "width": self.width,
                "height": self.height,
            },
            "page": self.page,
            "section": self.section,
            "para_id": self.para_id,
            "style_id": self.style_id,
            "metadata": self.metadata,
        }


@dataclass
class PageInfo:
    """
    페이지 정보
    
    Attributes:
        page_num (int): 페이지 번호
        width (float): 페이지 너비 (mm)
        height (float): 페이지 높이 (mm)
        margin_top (float): 상단 여백 (mm)
        margin_bottom (float): 하단 여백 (mm)
        margin_left (float): 좌측 여백 (mm)
        margin_right (float): 우측 여백 (mm)
    """
    page_num: int
    width: float
    height: float
    margin_top: float = 0
    margin_bottom: float = 0
    margin_left: float = 0
    margin_right: float = 0


def extract_layout_elements(doc: HwpxDocument) -> tuple[list[LayoutElement], list[PageInfo]]:
    """
    문서에서 레이아웃 요소들을 추출합니다.
    
    이 함수는 doc.to_json_with_layout()의 정보를 기반으로
    시각화에 필요한 필수 정보만 추출합니다.
    
    Args:
        doc: 파싱된 HWPX 문서
    
    Returns:
        tuple: (레이아웃 요소 리스트, 페이지 정보 리스트)
    
    사용 예시:
        doc = parse_hwpx_folder("results/hwpx_sample")
        elements, pages = extract_layout_elements(doc)
        
        for elem in elements:
            print(f"{elem.text[:20]}... at ({elem.x}, {elem.y})")
    """
    elements = []
    pages = []
    
    for section in doc.sections:
        # 페이지 정보 추출
        page_mm = section.page_props.to_mm()
        margin_mm = section.page_props.margin.to_mm()
        
        page_info = PageInfo(
            page_num=section.index,
            width=page_mm["width_mm"],
            height=page_mm["height_mm"],
            margin_top=margin_mm["top_mm"],
            margin_bottom=margin_mm["bottom_mm"],
            margin_left=margin_mm["left_mm"],
            margin_right=margin_mm["right_mm"],
        )
        pages.append(page_info)
        
        # 문단별 레이아웃 요소 추출
        for para in section.paragraphs:
            text = para.full_text
            if not text.strip() and not para.tables:
                continue
            
            # 바운딩 박스 계산
            bbox = para.get_bounding_box()
            
            if bbox and text.strip():
                elem = LayoutElement(
                    element_type="text",
                    text=text,
                    x=bbox["x_mm"],
                    y=bbox["y_mm"],
                    width=bbox["width_mm"],
                    height=bbox["height_mm"],
                    page=section.index,
                    section=section.index,
                    para_id=para.id,
                    style_id=para.style_id,
                    metadata={
                        "para_pr_id": para.para_pr_id,
                        "line_count": len(para.line_segments),
                    }
                )
                elements.append(elem)
            
            # 테이블 요소 추출
            for table in para.tables:
                table_size = table.size.to_mm()
                table_pos = table.position.to_mm()
                
                # 테이블 자체
                table_elem = LayoutElement(
                    element_type="table",
                    text=f"[Table {table.rows}×{table.cols}]",
                    x=table_pos["horz_offset_mm"],
                    y=table_pos["vert_offset_mm"],
                    width=table_size["width_mm"],
                    height=table_size["height_mm"],
                    page=section.index,
                    section=section.index,
                    metadata={
                        "rows": table.rows,
                        "cols": table.cols,
                        "id": table.id,
                        "z_order": table.z_order,
                    }
                )
                elements.append(table_elem)
                
                # 테이블 셀들
                for cell in table.cells:
                    cell_size = cell.size.to_mm()
                    cell_elem = LayoutElement(
                        element_type="table_cell",
                        text=cell.text,
                        x=table_pos["horz_offset_mm"],  # 셀별 정확한 위치 계산 필요
                        y=table_pos["vert_offset_mm"],
                        width=cell_size["width_mm"],
                        height=cell_size["height_mm"],
                        page=section.index,
                        section=section.index,
                        metadata={
                            "row": cell.row,
                            "col": cell.col,
                            "row_span": cell.row_span,
                            "col_span": cell.col_span,
                        }
                    )
                    elements.append(cell_elem)
    
    return elements, pages


def extract_layout_summary(doc: HwpxDocument) -> dict:
    """
    문서의 레이아웃 정보를 요약된 딕셔너리로 반환합니다.
    
    Args:
        doc: 파싱된 HWPX 문서
    
    Returns:
        dict: 레이아웃 요약 정보
    """
    elements, pages = extract_layout_elements(doc)
    
    return {
        "title": doc.title,
        "page_count": len(pages),
        "element_count": len(elements),
        "pages": [
            {
                "page_num": p.page_num,
                "size_mm": {"width": p.width, "height": p.height},
                "margins_mm": {
                    "top": p.margin_top,
                    "bottom": p.margin_bottom,
                    "left": p.margin_left,
                    "right": p.margin_right,
                }
            }
            for p in pages
        ],
        "elements": [e.to_dict() for e in elements],
    }


def visualize_document(
    doc: HwpxDocument,
    output_path: str | Path | None = None,
    page_num: int = 0,
    show_bbox: bool = True,
    show_text: bool = True,
    font_size: int = 8,
    figsize: tuple[float, float] | None = None,
    dpi: int = 100,
) -> Any:
    """
    문서를 화이트보드에 시각화합니다.
    
    바운딩 박스와 텍스트를 그려서 문서 레이아웃을 시각적으로 확인할 수 있습니다.
    
    Args:
        doc: 파싱된 HWPX 문서
        output_path: 이미지 저장 경로 (None이면 화면에 표시)
        page_num: 표시할 페이지 번호 (0부터 시작)
        show_bbox: 바운딩 박스 표시 여부
        show_text: 텍스트 표시 여부
        font_size: 텍스트 폰트 크기
        figsize: 그림 크기 (인치 단위, None이면 자동)
        dpi: 해상도
    
    Returns:
        matplotlib Figure 객체
    
    필요한 라이브러리:
        pip install matplotlib
    
    사용 예시:
        doc = parse_hwpx_folder("results/hwpx_sample")
        
        # 화면에 표시
        visualize_document(doc)
        
        # 파일로 저장
        visualize_document(doc, "output.png")
    """
    try:
        import matplotlib.pyplot as plt
        import matplotlib.patches as patches
        from matplotlib import font_manager
    except ImportError:
        raise ImportError(
            "matplotlib 라이브러리가 필요합니다.\n"
            "설치: pip install matplotlib"
        )
    
    # 레이아웃 요소 추출
    elements, pages = extract_layout_elements(doc)
    
    if page_num >= len(pages):
        raise ValueError(f"페이지 {page_num}이 존재하지 않습니다. (총 {len(pages)} 페이지)")
    
    page = pages[page_num]
    page_elements = [e for e in elements if e.page == page_num]
    
    # 그림 크기 설정 (A4 비율 유지)
    if figsize is None:
        scale = 0.5  # mm to inch 변환 (축소)
        figsize = (page.width * scale / 25.4, page.height * scale / 25.4)
    
    # Figure 생성
    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)
    
    # 배경 (화이트보드)
    ax.set_facecolor('white')
    ax.set_xlim(0, page.width)
    ax.set_ylim(page.height, 0)  # Y축 반전 (위에서 아래로)
    ax.set_aspect('equal')
    
    # 페이지 테두리
    page_rect = patches.Rectangle(
        (0, 0), page.width, page.height,
        linewidth=2, edgecolor='black', facecolor='white'
    )
    ax.add_patch(page_rect)
    
    # 여백 영역 표시 (연한 회색)
    margin_rect = patches.Rectangle(
        (page.margin_left, page.margin_top),
        page.width - page.margin_left - page.margin_right,
        page.height - page.margin_top - page.margin_bottom,
        linewidth=1, edgecolor='lightgray', facecolor='none',
        linestyle='--'
    )
    ax.add_patch(margin_rect)
    
    # 색상 정의
    colors = {
        "text": {"edge": "blue", "face": "lightblue", "alpha": 0.3},
        "table": {"edge": "green", "face": "lightgreen", "alpha": 0.3},
        "table_cell": {"edge": "orange", "face": "lightyellow", "alpha": 0.2},
    }
    
    # 요소들 그리기
    for elem in page_elements:
        color = colors.get(elem.element_type, colors["text"])
        
        # 좌표 보정 (여백 기준)
        x = page.margin_left + elem.x
        y = page.margin_top + elem.y
        
        if show_bbox:
            # 바운딩 박스
            rect = patches.Rectangle(
                (x, y), elem.width, elem.height,
                linewidth=1,
                edgecolor=color["edge"],
                facecolor=color["face"],
                alpha=color["alpha"]
            )
            ax.add_patch(rect)
        
        if show_text and elem.text.strip():
            # 텍스트 표시 (너무 긴 텍스트는 잘라서 표시)
            display_text = elem.text.strip()
            if len(display_text) > 30:
                display_text = display_text[:27] + "..."
            
            # 텍스트 위치 (박스 중앙 또는 왼쪽 상단)
            text_x = x + 1  # 약간의 패딩
            text_y = y + elem.height / 2
            
            ax.text(
                text_x, text_y,
                display_text,
                fontsize=font_size,
                verticalalignment='center',
                horizontalalignment='left',
                color='black',
                clip_on=True,
                fontfamily='sans-serif',
            )
    
    # 제목
    ax.set_title(
        f"{doc.title} - Page {page_num + 1}/{len(pages)}",
        fontsize=12,
        fontweight='bold'
    )
    
    # 축 레이블
    ax.set_xlabel("mm")
    ax.set_ylabel("mm")
    
    # 그리드 (선택적)
    ax.grid(True, linestyle=':', alpha=0.3)
    
    plt.tight_layout()
    
    # 저장 또는 표시
    if output_path:
        plt.savefig(output_path, dpi=dpi, bbox_inches='tight', facecolor='white')
        print(f"✅ 이미지 저장: {output_path}")
    
    return fig


def create_document_viewer(
    doc: HwpxDocument,
    output_dir: str | Path | None = None,
) -> list[Path]:
    """
    문서의 모든 페이지를 이미지로 저장하는 뷰어를 생성합니다.
    
    Args:
        doc: 파싱된 HWPX 문서
        output_dir: 출력 디렉토리 (None이면 현재 디렉토리)
    
    Returns:
        list[Path]: 생성된 이미지 파일 경로 리스트
    
    사용 예시:
        doc = parse_hwpx_folder("results/hwpx_sample")
        images = create_document_viewer(doc, "output_images")
    """
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        raise ImportError(
            "matplotlib 라이브러리가 필요합니다.\n"
            "설치: pip install matplotlib"
        )
    
    if output_dir is None:
        output_dir = Path(".")
    else:
        output_dir = Path(output_dir)
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    _, pages = extract_layout_elements(doc)
    saved_files = []
    
    print(f"📄 {doc.title} 문서 시각화 중...")
    
    for page_num in range(len(pages)):
        output_path = output_dir / f"{doc.title}_page_{page_num + 1:03d}.png"
        
        fig = visualize_document(
            doc,
            output_path=output_path,
            page_num=page_num,
            show_bbox=True,
            show_text=True,
        )
        plt.close(fig)  # 메모리 해제
        
        saved_files.append(output_path)
        print(f"  ✅ Page {page_num + 1}: {output_path}")
    
    print(f"\n📁 총 {len(saved_files)}개 이미지 생성 완료")
    
    return saved_files


def visualize_document_interactive(doc: HwpxDocument):
    """
    문서를 인터랙티브하게 시각화합니다 (Jupyter Notebook용).
    
    마우스를 올리면 해당 요소의 정보가 표시됩니다.
    
    Args:
        doc: 파싱된 HWPX 문서
    
    사용 예시 (Jupyter Notebook):
        doc = parse_hwpx_folder("results/hwpx_sample")
        visualize_document_interactive(doc)
    """
    try:
        import matplotlib.pyplot as plt
        from matplotlib.widgets import Slider
    except ImportError:
        raise ImportError(
            "matplotlib 라이브러리가 필요합니다.\n"
            "설치: pip install matplotlib"
        )
    
    elements, pages = extract_layout_elements(doc)
    
    if not pages:
        print("표시할 페이지가 없습니다.")
        return
    
    # 초기 페이지 표시
    fig = visualize_document(doc, page_num=0)
    
    # 슬라이더 추가 (여러 페이지인 경우)
    if len(pages) > 1:
        plt.subplots_adjust(bottom=0.15)
        ax_slider = plt.axes([0.2, 0.02, 0.6, 0.03])
        slider = Slider(
            ax_slider, 'Page',
            1, len(pages),
            valinit=1,
            valstep=1
        )
        
        def update(val):
            page_num = int(val) - 1
            plt.clf()
            visualize_document(doc, page_num=page_num)
            plt.draw()
        
        slider.on_changed(update)
    
    plt.show()


def visualize_document_pil(
    doc: HwpxDocument,
    output_path: str | Path,
    page_num: int = 0,
    show_bbox: bool = True,
    show_text: bool = True,
    scale: float = 3.0,
    font_size: int = 12,
) -> Path:
    """
    PIL을 사용하여 문서를 시각화합니다 (matplotlib 대안).
    
    matplotlib이 작동하지 않을 때 사용할 수 있는 대안입니다.
    
    Args:
        doc: 파싱된 HWPX 문서
        output_path: 이미지 저장 경로
        page_num: 표시할 페이지 번호
        show_bbox: 바운딩 박스 표시 여부
        show_text: 텍스트 표시 여부
        scale: 확대 비율 (1mm = scale 픽셀)
        font_size: 폰트 크기
    
    Returns:
        Path: 저장된 이미지 경로
    
    필요한 라이브러리:
        pip install Pillow
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        raise ImportError(
            "Pillow 라이브러리가 필요합니다.\n"
            "설치: pip install Pillow"
        )
    
    # 레이아웃 요소 추출
    elements, pages = extract_layout_elements(doc)
    
    if page_num >= len(pages):
        raise ValueError(f"페이지 {page_num}이 존재하지 않습니다.")
    
    page = pages[page_num]
    page_elements = [e for e in elements if e.page == page_num]
    
    # 이미지 크기 계산
    img_width = int(page.width * scale)
    img_height = int(page.height * scale)
    
    # 이미지 생성 (흰색 배경)
    img = Image.new('RGB', (img_width, img_height), 'white')
    draw = ImageDraw.Draw(img)
    
    # 기본 폰트 (시스템 폰트 사용 시도)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/AppleSDGothicNeo.ttc", font_size)
    except:
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", font_size)
        except:
            font = ImageFont.load_default()
    
    # 페이지 테두리
    draw.rectangle(
        [(0, 0), (img_width - 1, img_height - 1)],
        outline='black',
        width=2
    )
    
    # 여백 영역 표시
    margin_left = int(page.margin_left * scale)
    margin_top = int(page.margin_top * scale)
    margin_right = int((page.width - page.margin_right) * scale)
    margin_bottom = int((page.height - page.margin_bottom) * scale)
    draw.rectangle(
        [(margin_left, margin_top), (margin_right, margin_bottom)],
        outline='lightgray',
        width=1
    )
    
    # 색상 정의
    colors = {
        "text": {"outline": "blue", "fill": (173, 216, 230, 100)},
        "table": {"outline": "green", "fill": (144, 238, 144, 100)},
        "table_cell": {"outline": "orange", "fill": (255, 255, 224, 100)},
    }
    
    # 요소들 그리기
    for elem in page_elements:
        color = colors.get(elem.element_type, colors["text"])
        
        # 좌표 변환
        x1 = int((page.margin_left + elem.x) * scale)
        y1 = int((page.margin_top + elem.y) * scale)
        x2 = int((page.margin_left + elem.x + elem.width) * scale)
        y2 = int((page.margin_top + elem.y + elem.height) * scale)
        
        if show_bbox:
            # 바운딩 박스
            draw.rectangle(
                [(x1, y1), (x2, y2)],
                outline=color["outline"],
                width=1
            )
        
        if show_text and elem.text.strip():
            # 텍스트 표시
            display_text = elem.text.strip()
            if len(display_text) > 25:
                display_text = display_text[:22] + "..."
            
            # 텍스트가 박스 안에 들어가도록
            try:
                draw.text(
                    (x1 + 2, y1 + 2),
                    display_text,
                    fill='black',
                    font=font
                )
            except:
                pass  # 폰트 문제 시 무시
    
    # 제목 추가
    title = f"{doc.title} - Page {page_num + 1}/{len(pages)}"
    draw.text((10, 10), title, fill='black', font=font)
    
    # 저장
    output_path = Path(output_path)
    img.save(output_path)
    print(f"✅ 이미지 저장: {output_path}")
    
    return output_path


def create_document_viewer_pil(
    doc: HwpxDocument,
    output_dir: str | Path | None = None,
    scale: float = 3.0,
) -> list[Path]:
    """
    PIL을 사용하여 모든 페이지를 이미지로 저장합니다.
    
    Args:
        doc: 파싱된 HWPX 문서
        output_dir: 출력 디렉토리
        scale: 확대 비율
    
    Returns:
        list[Path]: 생성된 이미지 파일 경로 리스트
    """
    if output_dir is None:
        output_dir = Path(".")
    else:
        output_dir = Path(output_dir)
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    _, pages = extract_layout_elements(doc)
    saved_files = []
    
    print(f"📄 {doc.title} 문서 시각화 중 (PIL)...")
    
    for page_num in range(len(pages)):
        output_path = output_dir / f"{doc.title}_page_{page_num + 1:03d}.png"
        
        visualize_document_pil(
            doc,
            output_path=output_path,
            page_num=page_num,
            show_bbox=True,
            show_text=True,
            scale=scale,
        )
        
        saved_files.append(output_path)
    
    print(f"\n📁 총 {len(saved_files)}개 이미지 생성 완료")
    
    return saved_files


# 테스트용 메인 실행
if __name__ == "__main__" and len(sys.argv) > 1 and sys.argv[1] == "--visualize":
    # 시각화 테스트
    if len(sys.argv) > 2:
        folder = sys.argv[2]
    else:
        folder = "results/hwpx_sample"
    
    print(f"시각화 중: {folder}")
    
    doc = parse_hwpx_folder(folder)
    
    # 레이아웃 요약 출력
    summary = extract_layout_summary(doc)
    print(f"\n📊 레이아웃 요약:")
    print(f"  - 페이지 수: {summary['page_count']}")
    print(f"  - 요소 수: {summary['element_count']}")
    
    # 시각화
    try:
        output_path = Path(folder).parent / f"{doc.title}_visualization.png"
        visualize_document(doc, output_path=output_path)
        
        # JSON으로 레이아웃 정보 저장
        layout_json = Path(folder).parent / f"{doc.title}_layout_elements.json"
        with open(layout_json, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        print(f"✅ 레이아웃 JSON 저장: {layout_json}")
        
    except ImportError as e:
        print(f"❌ {e}")
