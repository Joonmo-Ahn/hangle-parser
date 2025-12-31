"""
HWPX Parser - HWPX 문서 파싱 및 레이아웃 정보 추출

HWPX 파일은 ZIP 압축된 XML 기반 문서입니다.
이 파서는 텍스트, 표, 레이아웃 정보를 추출하며,
특히 정확한 바운딩 박스 좌표를 제공하여 시각화가 가능합니다.

좌표 단위:
    HWPUNIT: 1 HWPUNIT = 1/7200 인치 ≈ 0.00353mm

사용 예시:
    from hwpx_parser import parse_hwpx, extract_layout_elements
    
    # 파싱
    doc = parse_hwpx("document.hwpx")

    # 텍스트 추출
    print(doc.to_text())

    # 레이아웃 정보 추출
    elements = extract_layout_elements(doc)
    for elem in elements:
        print(f"{elem.text[:20]}... at ({elem.x}, {elem.y})")
"""

from __future__ import annotations
import xml.etree.ElementTree as ET
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Iterator, Any
import json
import zipfile
import tempfile
import shutil


# =============================================================================
# 상수 및 네임스페이스
# =============================================================================

# XML 네임스페이스
NS = {
    "sec": "http://www.hancom.co.kr/hwpml/2011/section",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hv": "http://www.hancom.co.kr/hwpml/2011/version",
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
}

# HWPUNIT을 mm로 변환하는 상수
HWPUNIT_TO_MM = 25.4 / 7200  # ≈ 0.00353mm


# =============================================================================
# 레이아웃 데이터 클래스
# =============================================================================

@dataclass
class BoundingBox:
    """
    바운딩 박스 - 요소의 절대 좌표 (mm 단위)

    Attributes:
        x: X 좌표 (페이지 왼쪽 상단 기준)
        y: Y 좌표 (페이지 왼쪽 상단 기준)
        width: 너비
        height: 높이
    """
    x: float = 0.0
    y: float = 0.0
    width: float = 0.0
    height: float = 0.0

    def to_dict(self) -> dict:
        return {
            "x": round(self.x, 2),
            "y": round(self.y, 2),
            "width": round(self.width, 2),
            "height": round(self.height, 2),
            "x2": round(self.x + self.width, 2),
            "y2": round(self.y + self.height, 2),
        }

    def is_valid(self) -> bool:
        """유효한 좌표인지 확인 (모두 0이 아닌지)"""
        return not (self.x == 0 and self.y == 0 and self.width == 0 and self.height == 0)

    def __repr__(self):
        return f"BBox({self.x:.1f}, {self.y:.1f}, {self.width:.1f}×{self.height:.1f})"


@dataclass
class Position:
    """요소의 위치 정보 (HWPUNIT)"""
    vert_rel_to: str = ""
    horz_rel_to: str = ""
    vert_align: str = ""
    horz_align: str = ""
    vert_offset: int = 0
    horz_offset: int = 0
    treat_as_char: bool = False
    flow_with_text: bool = False

    def to_mm(self) -> dict:
        return {
            "vert_offset_mm": round(self.vert_offset * HWPUNIT_TO_MM, 2),
            "horz_offset_mm": round(self.horz_offset * HWPUNIT_TO_MM, 2),
        }


@dataclass
class Size:
    """요소의 크기 정보 (HWPUNIT)"""
    width: int = 0
    height: int = 0
    width_rel_to: str = "ABSOLUTE"
    height_rel_to: str = "ABSOLUTE"

    def to_mm(self) -> dict:
        return {
            "width_mm": round(self.width * HWPUNIT_TO_MM, 2),
            "height_mm": round(self.height * HWPUNIT_TO_MM, 2),
        }


@dataclass
class Margin:
    """여백 정보 (HWPUNIT)"""
    left: int = 0
    right: int = 0
    top: int = 0
    bottom: int = 0

    def to_mm(self) -> dict:
        return {
            "left_mm": round(self.left * HWPUNIT_TO_MM, 2),
            "right_mm": round(self.right * HWPUNIT_TO_MM, 2),
            "top_mm": round(self.top * HWPUNIT_TO_MM, 2),
            "bottom_mm": round(self.bottom * HWPUNIT_TO_MM, 2),
        }


@dataclass
class LineSegment:
    """
    텍스트 라인 세그먼트 - 한 줄의 레이아웃 정보

    이 정보가 실제 문서에서의 텍스트 위치를 나타냅니다.
    """
    text_pos: int = 0       # 텍스트 시작 위치 (문자 인덱스)
    vert_pos: int = 0       # 수직 위치 (HWPUNIT, 섹션 시작 기준)
    vert_size: int = 0      # 수직 크기/줄 높이
    text_height: int = 0    # 실제 텍스트 높이
    baseline: int = 0       # 베이스라인 위치
    spacing: int = 0        # 줄 간격
    horz_pos: int = 0       # 수평 위치 (HWPUNIT)
    horz_size: int = 0      # 수평 크기/줄 너비

    def to_mm(self) -> dict:
        return {
            "x_mm": round(self.horz_pos * HWPUNIT_TO_MM, 2),
            "y_mm": round(self.vert_pos * HWPUNIT_TO_MM, 2),
            "width_mm": round(self.horz_size * HWPUNIT_TO_MM, 2),
            "height_mm": round(self.vert_size * HWPUNIT_TO_MM, 2),
        }

    def to_bbox(self, margin_left: float = 0, margin_top: float = 0) -> BoundingBox:
        """BoundingBox로 변환 (여백 포함)"""
        return BoundingBox(
            x=margin_left + self.horz_pos * HWPUNIT_TO_MM,
            y=margin_top + self.vert_pos * HWPUNIT_TO_MM,
            width=self.horz_size * HWPUNIT_TO_MM,
            height=self.vert_size * HWPUNIT_TO_MM,
        )


@dataclass
class PageProperties:
    """페이지 속성"""
    width: int = 0
    height: int = 0
    landscape: str = "NARROWLY"  # WIDELY = 가로
    gutter_type: str = "LEFT_ONLY"
    margin: Margin = field(default_factory=Margin)

    def to_mm(self) -> dict:
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
    """테이블 셀"""
    row: int
    col: int
    text: str
    row_span: int = 1
    col_span: int = 1
    size: Size = field(default_factory=Size)
    margin: Margin = field(default_factory=Margin)
    border_fill_id: str = ""
    bbox: BoundingBox = field(default_factory=BoundingBox)


@dataclass
class Table:
    """테이블"""
    rows: int
    cols: int
    cells: list[TableCell] = field(default_factory=list)
    id: str = ""
    z_order: int = 0
    position: Position = field(default_factory=Position)
    size: Size = field(default_factory=Size)
    outer_margin: Margin = field(default_factory=Margin)
    inner_margin: Margin = field(default_factory=Margin)
    bbox: BoundingBox = field(default_factory=BoundingBox)

    def to_markdown(self) -> str:
        """마크다운 테이블로 변환"""
        if not self.cells:
            return ""

        grid = [["" for _ in range(self.cols)] for _ in range(self.rows)]
        for cell in self.cells:
            if 0 <= cell.row < self.rows and 0 <= cell.col < self.cols:
                grid[cell.row][cell.col] = cell.text.replace("|", "\\|").replace("\n", " ")

        lines = []
        for i, row in enumerate(grid):
            lines.append("| " + " | ".join(row) + " |")
            if i == 0:
                lines.append("|" + "|".join(["---"] * self.cols) + "|")

        return "\n".join(lines)

    def get_cell_bboxes(self) -> list[tuple[TableCell, BoundingBox]]:
        """
        각 셀의 바운딩 박스 계산 (개선 버전)
        
        개선 사항:
        1. 병합 셀(rowSpan/colSpan)을 고려한 너비/높이 계산
        2. 표 경계를 벗어나지 않도록 클리핑
        3. 각 셀의 실제 크기 정보 우선 사용
        """
        result = []

        if not self.bbox.is_valid():
            return result

        # 열 너비 배열 초기화 (기본값: 균등 분배)
        default_col_width = self.bbox.width / max(self.cols, 1)
        col_widths = [default_col_width] * self.cols
        
        # 각 열의 실제 너비 추출 (병합되지 않은 셀 기준)
        for cell in self.cells:
            if cell.col_span == 1 and cell.size.width > 0:
                width_mm = cell.size.to_mm()["width_mm"]
                if width_mm > 0 and cell.col < self.cols:
                    col_widths[cell.col] = width_mm

        # 행 높이 배열 초기화 (기본값: 균등 분배)
        default_row_height = self.bbox.height / max(self.rows, 1)
        row_heights = [default_row_height] * self.rows
        
        # 각 행의 실제 높이 추출 (병합되지 않은 셀 기준)
        for cell in self.cells:
            if cell.row_span == 1 and cell.size.height > 0:
                height_mm = cell.size.to_mm()["height_mm"]
                if height_mm > 0 and cell.row < self.rows:
                    row_heights[cell.row] = height_mm

        # 총 너비/높이가 표 크기를 초과하면 비율로 조정
        total_width = sum(col_widths)
        if total_width > 0 and abs(total_width - self.bbox.width) > 1:
            scale = self.bbox.width / total_width
            col_widths = [w * scale for w in col_widths]
            
        total_height = sum(row_heights)
        if total_height > 0 and abs(total_height - self.bbox.height) > 1:
            scale = self.bbox.height / total_height
            row_heights = [h * scale for h in row_heights]

        # 각 셀의 바운딩 박스 계산
        for cell in self.cells:
            # 시작 위치
            x = self.bbox.x + sum(col_widths[:cell.col])
            y = self.bbox.y + sum(row_heights[:cell.row])
            
            # 셀 크기 (병합 고려)
            end_col = min(cell.col + cell.col_span, self.cols)
            end_row = min(cell.row + cell.row_span, self.rows)
            width = sum(col_widths[cell.col:end_col])
            height = sum(row_heights[cell.row:end_row])
            
            # 표 경계를 벗어나지 않도록 클리핑
            max_x = self.bbox.x + self.bbox.width
            max_y = self.bbox.y + self.bbox.height
            
            if x + width > max_x:
                width = max(max_x - x, 0)
            if y + height > max_y:
                height = max(max_y - y, 0)

            cell_bbox = BoundingBox(x=x, y=y, width=width, height=height)
            result.append((cell, cell_bbox))

        return result


@dataclass
class TextRun:
    """텍스트 런 - 동일 서식 텍스트 조각"""
    text: str
    char_pr_id: str = ""
    start_pos: int = 0  # 문단 내 시작 위치
    end_pos: int = 0    # 문단 내 끝 위치


@dataclass
class Paragraph:
    """문단"""
    id: str
    texts: list[str] = field(default_factory=list)
    text_runs: list[TextRun] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    para_pr_id: str = ""
    style_id: str = ""
    line_segments: list[LineSegment] = field(default_factory=list)
    page_break: bool = False
    column_break: bool = False
    bbox: BoundingBox = field(default_factory=BoundingBox)

    @property
    def full_text(self) -> str:
        return "".join(self.texts)

    def calculate_bbox(self, margin_left: float = 0, margin_top: float = 0) -> BoundingBox:
        """라인 세그먼트에서 바운딩 박스 계산"""
        if not self.line_segments:
            return BoundingBox()

        # 유효한 라인 세그먼트 필터링 (모두 0이 아닌 것)
        valid_segments = [
            ls for ls in self.line_segments
            if ls.horz_size > 0 or ls.vert_size > 0
        ]

        if not valid_segments:
            return BoundingBox()

        # 전체 범위 계산
        min_x = min(ls.horz_pos for ls in valid_segments)
        max_x = max(ls.horz_pos + ls.horz_size for ls in valid_segments)
        min_y = min(ls.vert_pos for ls in valid_segments)
        max_y = max(ls.vert_pos + ls.vert_size for ls in valid_segments)

        return BoundingBox(
            x=margin_left + min_x * HWPUNIT_TO_MM,
            y=margin_top + min_y * HWPUNIT_TO_MM,
            width=(max_x - min_x) * HWPUNIT_TO_MM,
            height=(max_y - min_y) * HWPUNIT_TO_MM,
        )

    def get_char_bboxes(self, margin_left: float = 0, margin_top: float = 0) -> list[tuple[str, BoundingBox]]:
        """
        각 문자의 바운딩 박스 추정

        라인 세그먼트 정보를 기반으로 각 문자의 위치를 추정합니다.
        정확한 글자별 좌표가 필요한 경우 사용합니다.
        """
        result = []
        text = self.full_text

        if not text or not self.line_segments:
            return result

        # 각 라인 세그먼트에 해당하는 텍스트 매핑
        sorted_segments = sorted(self.line_segments, key=lambda x: x.text_pos)

        for i, seg in enumerate(sorted_segments):
            # 세그먼트의 텍스트 범위
            start = seg.text_pos
            if i + 1 < len(sorted_segments):
                end = sorted_segments[i + 1].text_pos
            else:
                end = len(text)

            seg_text = text[start:end]
            if not seg_text:
                continue

            # 각 문자의 너비 추정 (균등 분배)
            char_width = seg.horz_size * HWPUNIT_TO_MM / max(len(seg_text), 1)

            for j, char in enumerate(seg_text):
                char_bbox = BoundingBox(
                    x=margin_left + (seg.horz_pos * HWPUNIT_TO_MM) + (j * char_width),
                    y=margin_top + seg.vert_pos * HWPUNIT_TO_MM,
                    width=char_width,
                    height=seg.vert_size * HWPUNIT_TO_MM,
                )
                result.append((char, char_bbox))

        return result


@dataclass
class Section:
    """섹션"""
    index: int
    paragraphs: list[Paragraph] = field(default_factory=list)
    page_props: PageProperties = field(default_factory=PageProperties)

    @property
    def full_text(self) -> str:
        return "\n".join(p.full_text for p in self.paragraphs if p.full_text)


@dataclass
class VersionInfo:
    """버전 정보"""
    application: str = ""
    app_version: str = ""
    xml_version: str = ""


@dataclass
class HwpxDocument:
    """HWPX 문서"""
    folder_path: Path
    version: VersionInfo = field(default_factory=VersionInfo)
    sections: list[Section] = field(default_factory=list)
    metadata: dict = field(default_factory=dict)

    @property
    def title(self) -> str:
        return self.folder_path.stem

    def to_text(self) -> str:
        """전체 텍스트 추출"""
        return "\n\n".join(s.full_text for s in self.sections if s.full_text)

    def to_markdown(self) -> str:
        """마크다운으로 변환"""
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

    def to_json(self) -> str:
        """JSON으로 변환"""
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
        """레이아웃 정보 포함 JSON"""
        data = {
            "title": self.title,
            "version": asdict(self.version),
            "unit_info": {
                "description": "좌표 단위는 mm",
                "hwpunit_to_mm": HWPUNIT_TO_MM,
            },
            "sections": []
        }

        for section in self.sections:
            page_mm = section.page_props.to_mm()
            margin_mm = section.page_props.margin.to_mm()

            section_data = {
                "index": section.index,
                "page": {
                    "width_mm": page_mm["width_mm"],
                    "height_mm": page_mm["height_mm"],
                    "orientation": page_mm["orientation"],
                    "margins_mm": margin_mm,
                },
                "paragraphs": []
            }

            for para in section.paragraphs:
                if not para.full_text and not para.tables:
                    continue

                # 바운딩 박스 계산
                bbox = para.calculate_bbox(
                    margin_mm["left_mm"],
                    margin_mm["top_mm"]
                )

                para_data = {
                    "id": para.id,
                    "text": para.full_text,
                    "style_id": para.style_id,
                    "bbox": bbox.to_dict() if bbox.is_valid() else None,
                    "line_segments": [
                        {
                            "text_pos": ls.text_pos,
                            **ls.to_mm(),
                        }
                        for ls in para.line_segments
                    ],
                    "tables": []
                }

                for table in para.tables:
                    table_data = {
                        "id": table.id,
                        "rows": table.rows,
                        "cols": table.cols,
                        "bbox": table.bbox.to_dict() if table.bbox.is_valid() else None,
                        "cells": [
                            {
                                "row": c.row,
                                "col": c.col,
                                "text": c.text,
                                "row_span": c.row_span,
                                "col_span": c.col_span,
                                "bbox": c.bbox.to_dict() if c.bbox.is_valid() else None,
                            }
                            for c in table.cells
                        ]
                    }
                    para_data["tables"].append(table_data)

                section_data["paragraphs"].append(para_data)

            data["sections"].append(section_data)

        return json.dumps(data, ensure_ascii=False, indent=2)


# =============================================================================
# 파서 클래스
# =============================================================================

class HwpxParser:
    """HWPX 폴더 파서"""

    def __init__(self, folder_path: str | Path):
        self.folder_path = Path(folder_path)

        if not self.folder_path.exists():
            raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {folder_path}")

        self.contents_dir = self.folder_path / "Contents"
        if not self.contents_dir.exists():
            raise FileNotFoundError(f"Contents 폴더를 찾을 수 없습니다: {self.contents_dir}")

    def parse(self) -> HwpxDocument:
        """문서 전체 파싱"""
        doc = HwpxDocument(folder_path=self.folder_path)

        doc.version = self._parse_version()
        doc.metadata = self._parse_metadata()
        doc.sections = list(self._parse_sections())

        # 바운딩 박스 계산
        self._calculate_all_bboxes(doc)

        return doc

    def _parse_version(self) -> VersionInfo:
        """version.xml 파싱"""
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

            section.page_props = self._parse_page_properties(root)

            # 직접 자식 문단만 파싱 (테이블 내부 문단 제외)
            # root의 직접 자식 중 p 태그만 찾기
            for child in root:
                if self._strip_ns(child.tag) == "p":
                    para = self._parse_paragraph(child, is_table_cell=False)
                    if para.texts or para.tables:
                        section.paragraphs.append(para)
        except ET.ParseError as e:
            print(f"XML 파싱 오류 ({section_file}): {e}")

        return section

    def _parse_page_properties(self, root) -> PageProperties:
        """페이지 속성 추출"""
        props = PageProperties()

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

    def _parse_paragraph(self, p_elem, is_table_cell: bool = False) -> Paragraph:
        """
        문단 파싱

        Args:
            p_elem: 문단 XML 요소
            is_table_cell: 테이블 셀 내부 문단인지 여부
        """
        para = Paragraph(
            id=p_elem.get("id", ""),
            para_pr_id=p_elem.get("paraPrIDRef", ""),
            style_id=p_elem.get("styleIDRef", ""),
            page_break=p_elem.get("pageBreak", "0") == "1",
            column_break=p_elem.get("columnBreak", "0") == "1",
        )

        text_pos = 0

        # iter() 대신 직접 자식만 순회하여 중첩 문단 방지
        def process_element(elem, depth=0):
            nonlocal text_pos

            tag = self._strip_ns(elem.tag)

            # 텍스트 런
            if tag == "run":
                char_pr_id = elem.get("charPrIDRef", "")
                for child in elem:
                    child_tag = self._strip_ns(child.tag)
                    if child_tag == "t" and child.text:
                        text = child.text
                        para.texts.append(text)
                        para.text_runs.append(TextRun(
                            text=text,
                            char_pr_id=char_pr_id,
                            start_pos=text_pos,
                            end_pos=text_pos + len(text)
                        ))
                        text_pos += len(text)
                    # 테이블 처리 (run 내부에 있을 수 있음)
                    elif child_tag == "tbl" and not is_table_cell:
                        table = self._parse_table(child)
                        if table:
                            para.tables.append(table)

            # 라인 세그먼트 배열
            elif tag == "linesegarray":
                for child in elem:
                    if self._strip_ns(child.tag) == "lineseg":
                        ls = LineSegment(
                            text_pos=int(child.get("textpos", 0)),
                            vert_pos=int(child.get("vertpos", 0)),
                            vert_size=int(child.get("vertsize", 0)),
                            text_height=int(child.get("textheight", 0)),
                            baseline=int(child.get("baseline", 0)),
                            spacing=int(child.get("spacing", 0)),
                            horz_pos=int(child.get("horzpos", 0)),
                            horz_size=int(child.get("horzsize", 0)),
                        )
                        para.line_segments.append(ls)

            # 개별 라인 세그먼트 (linesegarray 없이 직접 있는 경우)
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

            # 테이블 (직접 자식인 경우)
            elif tag == "tbl" and not is_table_cell:
                table = self._parse_table(elem)
                if table:
                    para.tables.append(table)

        # 직접 자식 요소만 처리
        for child in p_elem:
            process_element(child)

        return para

    def _parse_table(self, tbl_elem) -> Table | None:
        """테이블 파싱"""
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

        # 테이블 레이아웃 정보
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

    def _calculate_all_bboxes(self, doc: HwpxDocument):
        """
        모든 요소의 바운딩 박스 계산 (개선 버전)
        
        개선 사항:
        1. 상대 위치 기준점(vertRelTo, horzRelTo) 적용
        2. 페이지 경계 검증
        3. 이전 문단 위치를 기반으로 한 누적 Y 좌표 계산
        """
        for section in doc.sections:
            margin_mm = section.page_props.margin.to_mm()
            margin_left = margin_mm["left_mm"]
            margin_top = margin_mm["top_mm"]
            margin_right = margin_mm["right_mm"]
            margin_bottom = margin_mm["bottom_mm"]
            
            page_mm = section.page_props.to_mm()
            page_width = page_mm["width_mm"]
            page_height = page_mm["height_mm"]
            
            # 콘텐츠 영역
            content_width = page_width - margin_left - margin_right
            content_height = page_height - margin_top - margin_bottom
            
            # 현재 Y 위치 추적 (문단 흐름용)
            current_y = margin_top

            for para in section.paragraphs:
                # 문단 바운딩 박스
                para.bbox = para.calculate_bbox(margin_left, margin_top)
                
                # 문단 바운딩 박스가 유효하면 현재 Y 위치 업데이트
                if para.bbox.is_valid():
                    current_y = para.bbox.y + para.bbox.height

                # 테이블 바운딩 박스
                for table in para.tables:
                    size_mm = table.size.to_mm()
                    pos_mm = table.position.to_mm()
                    
                    # 수평 위치 계산 (horzRelTo 기준)
                    horz_rel = table.position.horz_rel_to
                    if horz_rel == "PAGE":
                        table_x = pos_mm["horz_offset_mm"]
                    elif horz_rel == "COLUMN" or horz_rel == "PARA":
                        table_x = margin_left + pos_mm["horz_offset_mm"]
                    else:
                        # 기본값: 마진 기준
                        table_x = margin_left + pos_mm["horz_offset_mm"]
                    
                    # 수직 위치 계산 (vertRelTo 기준)
                    vert_rel = table.position.vert_rel_to
                    if vert_rel == "PAGE":
                        table_y = pos_mm["vert_offset_mm"]
                    elif vert_rel == "PARA":
                        # 문단 기준: 문단의 바운딩 박스 Y + 오프셋
                        if para.bbox.is_valid():
                            table_y = para.bbox.y + pos_mm["vert_offset_mm"]
                        else:
                            table_y = current_y + pos_mm["vert_offset_mm"]
                    elif table.position.treat_as_char:
                        # 문자처럼 취급: 현재 문단 위치
                        table_y = para.bbox.y if para.bbox.is_valid() else current_y
                    else:
                        # 기본값: 마진 기준
                        table_y = margin_top + pos_mm["vert_offset_mm"]
                    
                    # 페이지 경계 내로 클리핑
                    table_x = max(0, min(table_x, page_width - size_mm["width_mm"]))
                    table_y = max(0, min(table_y, page_height - size_mm["height_mm"]))
                    
                    # 테이블 너비가 페이지를 초과하면 조정
                    table_width = min(size_mm["width_mm"], page_width - table_x)
                    table_height = min(size_mm["height_mm"], page_height - table_y)

                    table.bbox = BoundingBox(
                        x=table_x,
                        y=table_y,
                        width=table_width,
                        height=table_height,
                    )

                    # 셀 바운딩 박스
                    cell_bboxes = table.get_cell_bboxes()
                    for cell, cell_bbox in cell_bboxes:
                        cell.bbox = cell_bbox
                    
                    # 테이블 후 현재 Y 위치 업데이트
                    current_y = max(current_y, table_y + table_height)

    @staticmethod
    def _strip_ns(tag: str) -> str:
        """XML 태그에서 네임스페이스 제거"""
        if "}" in tag:
            return tag.split("}")[-1]
        return tag


# =============================================================================
# 레이아웃 추출 함수
# =============================================================================

@dataclass
class LayoutElement:
    """
    레이아웃 요소 - 시각화/추출용

    Attributes:
        element_type: 요소 유형 (text, table, table_cell)
        text: 텍스트 내용
        x, y, width, height: 좌표 (mm)
        page: 페이지 번호
        section: 섹션 번호
        para_id: 문단 ID
        style_id: 스타일 ID
        metadata: 추가 정보
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
        return {
            "type": self.element_type,
            "text": self.text,
            "bbox": {
                "x": round(self.x, 2),
                "y": round(self.y, 2),
                "width": round(self.width, 2),
                "height": round(self.height, 2),
            },
            "page": self.page,
            "section": self.section,
            "para_id": self.para_id,
            "style_id": self.style_id,
            "metadata": self.metadata,
        }


@dataclass
class PageInfo:
    """페이지 정보"""
    page_num: int
    width: float
    height: float
    margin_top: float = 0
    margin_bottom: float = 0
    margin_left: float = 0
    margin_right: float = 0


def extract_layout_elements(doc: HwpxDocument) -> tuple[list[LayoutElement], list[PageInfo]]:
    """
    문서에서 레이아웃 요소를 추출합니다.

    Args:
        doc: 파싱된 HWPX 문서

    Returns:
        tuple: (레이아웃 요소 리스트, 페이지 정보 리스트)

    사용 예시:
        doc = parse_hwpx("document.hwpx")
        elements, pages = extract_layout_elements(doc)

        for elem in elements:
            print(f"{elem.text[:20]}... at ({elem.x:.1f}, {elem.y:.1f})")
    """
    elements = []
    pages = []

    for section in doc.sections:
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

        for para in section.paragraphs:
            text = para.full_text
            if not text.strip() and not para.tables:
                continue

            # 문단 요소
            if text.strip() and para.bbox.is_valid():
                elem = LayoutElement(
                    element_type="text",
                    text=text,
                    x=para.bbox.x,
                    y=para.bbox.y,
                    width=para.bbox.width,
                    height=para.bbox.height,
                    page=section.index,
                    section=section.index,
                    para_id=para.id,
                    style_id=para.style_id,
                    metadata={
                        "line_count": len(para.line_segments),
                    }
                )
                elements.append(elem)

            # 테이블 요소
            for table in para.tables:
                if table.bbox.is_valid():
                    table_elem = LayoutElement(
                        element_type="table",
                        text=f"[Table {table.rows}×{table.cols}]",
                        x=table.bbox.x,
                        y=table.bbox.y,
                        width=table.bbox.width,
                        height=table.bbox.height,
                        page=section.index,
                        section=section.index,
                        metadata={
                            "rows": table.rows,
                            "cols": table.cols,
                            "id": table.id,
                        }
                    )
                    elements.append(table_elem)

                    # 셀 요소
                    for cell in table.cells:
                        if cell.bbox.is_valid():
                            cell_elem = LayoutElement(
                                element_type="table_cell",
                                text=cell.text,
                                x=cell.bbox.x,
                                y=cell.bbox.y,
                                width=cell.bbox.width,
                                height=cell.bbox.height,
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
    """문서 레이아웃 요약"""
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


# =============================================================================
# 편의 함수
# =============================================================================

def parse_hwpx_folder(folder_path: str | Path) -> HwpxDocument:
    """HWPX 폴더 파싱"""
    parser = HwpxParser(folder_path)
    return parser.parse()


def parse_hwpx(file_path: str | Path, extract_dir: str | Path | None = None) -> HwpxDocument:
    """
    HWPX 파일 파싱

    Args:
        file_path: HWPX 파일 경로
        extract_dir: 압축 해제 디렉토리 (None이면 임시 디렉토리)

    Returns:
        HwpxDocument: 파싱된 문서
    """
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    if not file_path.suffix.lower() == ".hwpx":
        raise ValueError(f"HWPX 파일이 아닙니다: {file_path}")

    # 압축 해제
    if extract_dir is None:
        temp_dir = tempfile.mkdtemp(prefix="hwpx_")
        extract_path = Path(temp_dir) / file_path.stem
        cleanup = True
    else:
        extract_path = Path(extract_dir) / file_path.stem
        cleanup = False

    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            zf.extractall(extract_path)

        doc = parse_hwpx_folder(extract_path)
        doc.folder_path = file_path  # 원본 파일 경로 저장

        return doc

    finally:
        if cleanup and extract_path.exists():
            shutil.rmtree(extract_path.parent, ignore_errors=True)


# =============================================================================
# 메인
# =============================================================================

if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        print("사용법: python hwpx_parser.py <hwpx_file>")
        sys.exit(1)

    print(f"파싱 중: {file_path}")
    print("=" * 60)

    doc = parse_hwpx(file_path)

    print(f"\n📄 문서: {doc.title}")
    print(f"📋 버전: {doc.version.application} {doc.version.app_version}")
    print(f"📑 섹션 수: {len(doc.sections)}")

    for section in doc.sections:
        print(f"\n--- Section {section.index + 1} ---")
        page_mm = section.page_props.to_mm()
        print(f"  페이지: {page_mm['width_mm']}mm × {page_mm['height_mm']}mm")
        print(f"  문단 수: {len(section.paragraphs)}")

    # 레이아웃 요소 출력
    elements, pages = extract_layout_elements(doc)
    print(f"\n📐 레이아웃 요소: {len(elements)}개")

    for elem in elements[:5]:
        print(f"  - {elem.element_type}: ({elem.x:.1f}, {elem.y:.1f}) {elem.width:.1f}×{elem.height:.1f}mm")
        if elem.text:
            print(f"    텍스트: {elem.text[:50]}...")
