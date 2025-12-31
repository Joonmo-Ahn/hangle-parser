"""
Document Extractor - HWPX/HWP 문서에서 LLM/RAG에 적합한 구조화된 정보 추출

주요 기능:
1. 정확한 바운딩 박스 좌표 추출 (페이지 내 절대 좌표)
2. 표의 구조화 (제목/헤더/내용 분리)
3. 문서 요소의 계층적 구조화 (큰 제목 > 작은 제목 > 내용)
4. LLM에 적합한 구조화된 텍스트 출력

사용 예시:
    from document_extractor import extract_document, create_visualization_report
    from hwpx_parser import parse_hwpx
    from hwp_parser import parse_hwp

    # HWPX 문서
    doc = parse_hwpx("document.hwpx")
    extracted = extract_document(doc)

    # 구조화된 텍스트 (LLM용)
    print(extracted.to_structured_text())

    # 시각화 리포트 생성
    create_visualization_report(extracted, "output_dir")
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Literal, Optional
from pathlib import Path
import json
import re

# 이미지 추출 모듈 (선택적)
try:
    from image_extractor import (
        extract_images_from_hwp, 
        extract_images_from_hwpx,
        save_images_for_ocr,
        generate_images_report,
        EmbeddedImage,
    )
    HAS_IMAGE_EXTRACTOR = True
except ImportError:
    HAS_IMAGE_EXTRACTOR = False


# =============================================================================
# 데이터 클래스 정의
# =============================================================================

@dataclass
class BoundingBox:
    """바운딩 박스 (mm 단위)"""
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
        return not (self.x == 0 and self.y == 0 and self.width == 0 and self.height == 0)

    def __repr__(self):
        return f"BBox({self.x:.1f}, {self.y:.1f}, {self.width:.1f}×{self.height:.1f})"


@dataclass
class DocumentElement:
    """
    문서 요소

    Attributes:
        element_id: 요소 ID
        element_type: 요소 유형 (heading, paragraph, table, table_cell, image)
        text: 텍스트 내용
        bbox: 바운딩 박스
        page: 페이지 번호 (0부터)
        level: 제목 레벨 (1=큰 제목, 2=중간 제목, 3=작은 제목)
        parent_id: 부모 요소 ID
        children: 자식 요소 ID 리스트
        style: 스타일 정보
        metadata: 추가 메타데이터
    """
    element_id: str
    element_type: Literal["heading", "paragraph", "table", "table_cell", "image"]
    text: str
    bbox: BoundingBox
    page: int = 0
    level: int = 0
    parent_id: str = ""
    children: list[str] = field(default_factory=list)
    style: dict = field(default_factory=dict)
    metadata: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "id": self.element_id,
            "type": self.element_type,
            "text": self.text,
            "bbox": self.bbox.to_dict() if self.bbox.is_valid() else None,
            "page": self.page,
            "level": self.level,
            "parent_id": self.parent_id,
            "children": self.children,
            "style": self.style,
            "metadata": self.metadata,
        }


@dataclass
class TableStructure:
    """
    구조화된 표 정보

    표의 문맥을 파악하여 제목, 헤더, 데이터를 분리합니다.
    LLM이 표의 내용을 이해하기 쉽게 구조화합니다.
    """
    table_id: str
    title: str                          # 표 제목 (표 위의 텍스트)
    headers: list[list[str]]            # 헤더 행들
    rows: list[list[str]]               # 데이터 행들
    bbox: BoundingBox
    page: int = 0
    row_count: int = 0
    col_count: int = 0
    context: str = ""                   # 표 주변 문맥 (앞뒤 문단)
    metadata: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "id": self.table_id,
            "title": self.title,
            "headers": self.headers,
            "rows": self.rows,
            "bbox": self.bbox.to_dict() if self.bbox.is_valid() else None,
            "page": self.page,
            "row_count": self.row_count,
            "col_count": self.col_count,
            "context": self.context,
            "metadata": self.metadata,
        }

    def to_markdown(self) -> str:
        """마크다운 테이블로 변환"""
        lines = []
        if self.title:
            lines.append(f"**{self.title}**")
            lines.append("")

        all_rows = self.headers + self.rows
        if not all_rows:
            return ""

        max_cols = max(len(row) for row in all_rows) if all_rows else 0

        for i, row in enumerate(all_rows):
            padded_row = row + [""] * (max_cols - len(row))
            lines.append("| " + " | ".join(
                cell.replace("|", "\\|").replace("\n", " ")
                for cell in padded_row
            ) + " |")

            if i == len(self.headers) - 1 and self.headers:
                lines.append("|" + "|".join(["---"] * max_cols) + "|")

        return "\n".join(lines)

    def to_structured_text(self) -> str:
        """LLM에 적합한 구조화된 텍스트"""
        lines = []

        if self.title:
            lines.append(f"[표 제목] {self.title}")

        if self.headers:
            for i, header_row in enumerate(self.headers):
                header_text = " | ".join(header_row)
                if i == 0:
                    lines.append(f"[헤더] {header_text}")
                else:
                    lines.append(f"[부헤더] {header_text}")

        for i, row in enumerate(self.rows):
            row_text = " | ".join(row)
            lines.append(f"[행 {i+1}] {row_text}")

        return "\n".join(lines)


@dataclass
class HierarchicalSection:
    """
    계층적 섹션

    문서의 구조를 큰 제목 > 작은 제목 > 내용으로 계층화합니다.
    """
    section_id: str
    title: str
    level: int                          # 1=큰 제목, 2=중간 제목, 3=작은 제목
    content: list[str]                  # 해당 섹션의 내용 (문단들)
    tables: list[TableStructure]        # 해당 섹션의 표들
    children: list["HierarchicalSection"] = field(default_factory=list)
    bbox: BoundingBox = field(default_factory=BoundingBox)
    page: int = 0

    def to_dict(self) -> dict:
        return {
            "id": self.section_id,
            "title": self.title,
            "level": self.level,
            "content": self.content,
            "tables": [t.to_dict() for t in self.tables],
            "children": [c.to_dict() for c in self.children],
            "page": self.page,
        }

    def to_structured_text(self, indent: int = 0) -> str:
        """구조화된 텍스트로 변환"""
        lines = []
        prefix = "  " * indent

        # 제목
        level_marker = "#" * (self.level + 1)
        lines.append(f"{prefix}{level_marker} {self.title}")
        lines.append("")

        # 내용
        for content in self.content:
            if content.strip():
                lines.append(f"{prefix}{content}")
                lines.append("")

        # 표
        for table in self.tables:
            lines.append(f"{prefix}{table.to_structured_text()}")
            lines.append("")

        # 하위 섹션
        for child in self.children:
            lines.append(child.to_structured_text(indent))

        return "\n".join(lines)


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

    def to_dict(self) -> dict:
        return {
            "page_num": self.page_num,
            "width": round(self.width, 2),
            "height": round(self.height, 2),
            "margins": {
                "top": round(self.margin_top, 2),
                "bottom": round(self.margin_bottom, 2),
                "left": round(self.margin_left, 2),
                "right": round(self.margin_right, 2),
            }
        }


@dataclass
class ImageElement:
    """
    이미지 요소 (외부 OCR 연동용)
    
    Attributes:
        image_id: 이미지 ID (BIN0001 등)
        filename: 파일명
        format: 이미지 형식 (jpg, png, bmp 등)
        bbox: 바운딩 박스 (mm 단위)
        page: 페이지 번호 (0부터)
        pixel_width: 픽셀 너비
        pixel_height: 픽셀 높이
        file_size: 파일 크기 (bytes)
        saved_path: 저장된 파일 경로
        ocr_text: 외부 OCR 결과 텍스트 (연동 시 채움)
    """
    image_id: str
    filename: str
    format: str
    bbox: BoundingBox
    page: int = 0
    pixel_width: int = 0
    pixel_height: int = 0
    file_size: int = 0
    saved_path: str = ""
    ocr_text: str = ""  # 외부 OCR 결과
    
    def to_dict(self) -> dict:
        return {
            "image_id": self.image_id,
            "filename": self.filename,
            "format": self.format,
            "bbox": self.bbox.to_dict() if self.bbox.is_valid() else None,
            "page": self.page,
            "pixel_width": self.pixel_width,
            "pixel_height": self.pixel_height,
            "file_size": self.file_size,
            "saved_path": self.saved_path,
            "ocr_text": self.ocr_text,
        }
    
    def to_ocr_dict(self) -> dict:
        """외부 OCR 연동용 딕셔너리"""
        return {
            "image_id": self.image_id,
            "filename": self.filename,
            "format": self.format,
            "class": "image",
            "bbox_mm": {
                "x": round(self.bbox.x, 2),
                "y": round(self.bbox.y, 2),
                "width": round(self.bbox.width, 2),
                "height": round(self.bbox.height, 2),
            },
            "bbox_px": {
                "width": self.pixel_width,
                "height": self.pixel_height,
            },
            "page": self.page,
            "saved_path": self.saved_path,
            "ocr_text": self.ocr_text,
            "ocr_confidence": 0.0,
        }


@dataclass
class ExtractedDocument:
    """
    추출된 문서

    LLM/RAG에 사용하기 적합한 구조화된 문서 정보
    """
    title: str
    source_file: str
    file_type: str                      # "hwpx" or "hwp"
    pages: list[PageInfo]
    elements: list[DocumentElement]
    tables: list[TableStructure]
    headings: list[DocumentElement]
    paragraphs: list[DocumentElement]
    images: list[ImageElement] = field(default_factory=list)  # 이미지 요소
    hierarchical_sections: list[HierarchicalSection] = field(default_factory=list)
    metadata: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "title": self.title,
            "source_file": self.source_file,
            "file_type": self.file_type,
            "page_count": len(self.pages),
            "element_count": len(self.elements),
            "table_count": len(self.tables),
            "heading_count": len(self.headings),
            "image_count": len(self.images),
            "pages": [p.to_dict() for p in self.pages],
            "elements": [e.to_dict() for e in self.elements],
            "tables": [t.to_dict() for t in self.tables],
            "headings": [h.to_dict() for h in self.headings],
            "images": [i.to_dict() for i in self.images],
            "hierarchical_sections": [s.to_dict() for s in self.hierarchical_sections],
            "metadata": self.metadata,
        }

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)

    def to_structured_text(self) -> str:
        """
        LLM에 적합한 구조화된 텍스트로 변환

        문서의 계층 구조를 보존하면서 텍스트로 변환합니다.
        RAG 시스템에서 청킹하기 좋은 형태입니다.
        """
        lines = [f"# {self.title}", ""]

        # 메타 정보
        lines.append(f"[문서 유형] {self.file_type.upper()}")
        lines.append(f"[페이지 수] {len(self.pages)}")
        lines.append("")

        # 계층적 구조가 있으면 사용
        if self.hierarchical_sections:
            lines.append("## 문서 내용")
            lines.append("")
            for section in self.hierarchical_sections:
                lines.append(section.to_structured_text())
        else:
            # 없으면 순차적으로 출력
            lines.append("## 본문")
            lines.append("")

            current_page = -1
            for elem in self.elements:
                if elem.page != current_page:
                    current_page = elem.page
                    lines.append(f"### 페이지 {current_page + 1}")
                    lines.append("")

                if elem.element_type == "heading":
                    level = elem.level if elem.level > 0 else 1
                    lines.append("#" * (level + 2) + " " + elem.text)
                    lines.append("")
                elif elem.element_type == "paragraph":
                    if elem.text.strip():
                        lines.append(elem.text.strip())
                        lines.append("")

        # 표 목록
        if self.tables:
            lines.append("## 표 목록")
            lines.append("")
            for i, table in enumerate(self.tables):
                lines.append(f"### 표 {i + 1}")
                lines.append(table.to_structured_text())
                lines.append("")

        return "\n".join(lines)

    def to_rag_chunks(self, max_chunk_size: int = 1000) -> list[dict]:
        """
        RAG에 적합한 청크로 분할

        Args:
            max_chunk_size: 청크 최대 크기 (문자 수)

        Returns:
            list[dict]: 청크 리스트 (text, metadata 포함)
        """
        chunks = []

        # 제목별로 청킹
        for section in self.hierarchical_sections:
            chunk_text = section.to_structured_text()

            if len(chunk_text) <= max_chunk_size:
                chunks.append({
                    "text": chunk_text,
                    "metadata": {
                        "title": section.title,
                        "level": section.level,
                        "page": section.page,
                        "source": self.source_file,
                    }
                })
            else:
                # 큰 섹션은 문단별로 분할
                current_chunk = f"## {section.title}\n\n"
                for content in section.content:
                    if len(current_chunk) + len(content) > max_chunk_size:
                        if current_chunk.strip():
                            chunks.append({
                                "text": current_chunk.strip(),
                                "metadata": {
                                    "title": section.title,
                                    "level": section.level,
                                    "page": section.page,
                                    "source": self.source_file,
                                }
                            })
                        current_chunk = f"## {section.title} (계속)\n\n{content}\n\n"
                    else:
                        current_chunk += content + "\n\n"

                if current_chunk.strip():
                    chunks.append({
                        "text": current_chunk.strip(),
                        "metadata": {
                            "title": section.title,
                            "level": section.level,
                            "page": section.page,
                            "source": self.source_file,
                        }
                    })

        # 표도 별도 청크로
        for table in self.tables:
            chunks.append({
                "text": table.to_structured_text(),
                "metadata": {
                    "type": "table",
                    "title": table.title,
                    "page": table.page,
                    "source": self.source_file,
                }
            })

        return chunks

    def get_full_text(self) -> str:
        """전체 텍스트 추출"""
        texts = []
        for elem in self.elements:
            if elem.text.strip():
                texts.append(elem.text.strip())
        return "\n".join(texts)


# =============================================================================
# 제목 판별 함수
# =============================================================================

def is_heading(text: str, style_id: str = "") -> tuple[bool, int]:
    """
    텍스트가 제목인지 판별

    Args:
        text: 텍스트
        style_id: 스타일 ID

    Returns:
        tuple: (제목 여부, 레벨)
    """
    text = text.strip()
    if not text:
        return False, 0

    # 스타일 ID 기반 판별
    if style_id:
        try:
            style_num = int(style_id)
            if 1 <= style_num <= 6:
                return True, style_num
        except:
            pass

    # 패턴 기반 판별
    patterns = [
        # 레벨 1: 대제목
        (r'^제\s*\d+\s*장\s', 1),
        (r'^제\s*\d+\s*편\s', 1),
        (r'^Ⅰ\.\s', 1),
        (r'^Ⅱ\.\s', 1),
        (r'^Ⅲ\.\s', 1),
        (r'^Ⅳ\.\s', 1),
        (r'^Ⅴ\.\s', 1),

        # 레벨 2: 중제목
        (r'^제\s*\d+\s*조\s', 2),
        (r'^제\s*\d+\s*절\s', 2),
        (r'^\d+\.\s+[가-힣]', 2),
        (r'^[가-다]\.\s', 2),
        (r'^[1-9]\)\s', 2),
        (r'^【.+】', 2),
        (r'^\[.+\]$', 2),

        # 레벨 3: 소제목
        (r'^[가-힣]\)\s', 3),
        (r'^[ㄱ-ㅎ]\.\s', 3),
        (r'^①\s', 3),
        (r'^②\s', 3),
        (r'^③\s', 3),
        (r'^-\s+[가-힣]', 3),
    ]

    for pattern, level in patterns:
        if re.match(pattern, text):
            return True, level

    # 짧고 마침표로 끝나지 않는 텍스트
    if len(text) < 50 and not text.endswith(('.', '다', '요', '음', '임')):
        # 숫자나 특수문자로 시작하면 제목일 가능성
        if re.match(r'^[\d가-힣【\[①②③]', text):
            return True, 2

    return False, 0


def extract_table_title(prev_text: str) -> str:
    """
    표 제목 추출

    표 바로 앞의 텍스트에서 제목 추출
    """
    if not prev_text:
        return ""

    prev_text = prev_text.strip()

    # 너무 긴 텍스트는 제목이 아님
    if len(prev_text) > 100:
        return ""

    # 표 제목 패턴
    patterns = [
        r'^<표\s*\d*>\s*(.+)',
        r'^표\s*\d+[.:]\s*(.+)',
        r'^\[표\s*\d*\]\s*(.+)',
        r'^【표\s*\d*】\s*(.+)',
    ]

    for pattern in patterns:
        match = re.match(pattern, prev_text)
        if match:
            return match.group(1).strip()

    # 짧고 마침표로 끝나지 않으면 제목으로 간주
    if len(prev_text) < 50 and not prev_text.endswith('.'):
        return prev_text

    return ""


# =============================================================================
# HWPX 문서 추출
# =============================================================================

HWPUNIT_TO_MM = 25.4 / 7200


def extract_from_hwpx(doc) -> ExtractedDocument:
    """
    HWPX 문서에서 구조화된 정보 추출

    Args:
        doc: HwpxDocument 객체

    Returns:
        ExtractedDocument
    """
    elements = []
    tables = []
    headings = []
    paragraphs = []
    pages = []
    hierarchical_sections = []

    element_counter = 0
    section_counter = 0

    def next_elem_id() -> str:
        nonlocal element_counter
        element_counter += 1
        return f"elem_{element_counter:04d}"

    def next_section_id() -> str:
        nonlocal section_counter
        section_counter += 1
        return f"sec_{section_counter:04d}"

    # 현재 계층 구조 추적
    current_sections = {1: None, 2: None, 3: None}
    root_sections = []

    for section in doc.sections:
        # 페이지 정보
        page_mm = section.page_props.to_mm()
        margin_mm = section.page_props.margin.to_mm()
        page_height = page_mm["height_mm"]

        # 첫 페이지 생성
        current_page_num = len(pages)
        page_info = PageInfo(
            page_num=current_page_num,
            width=page_mm["width_mm"],
            height=page_height,
            margin_top=margin_mm["top_mm"],
            margin_bottom=margin_mm["bottom_mm"],
            margin_left=margin_mm["left_mm"],
            margin_right=margin_mm["right_mm"],
        )
        pages.append(page_info)

        margin_left = margin_mm["left_mm"]
        margin_top = margin_mm["top_mm"]
        prev_para_text = ""

        for para in section.paragraphs:
            # pageBreak 체크 - 새 페이지 시작
            if para.page_break:
                current_page_num = len(pages)
                new_page_info = PageInfo(
                    page_num=current_page_num,
                    width=page_mm["width_mm"],
                    height=page_height,
                    margin_top=margin_mm["top_mm"],
                    margin_bottom=margin_mm["bottom_mm"],
                    margin_left=margin_mm["left_mm"],
                    margin_right=margin_mm["right_mm"],
                )
                pages.append(new_page_info)
            text = para.full_text.strip()

            # 바운딩 박스 계산
            bbox = para.calculate_bbox(margin_left, margin_top)
            if not bbox.is_valid():
                bbox = BoundingBox(
                    x=margin_left,
                    y=margin_top,
                    width=page_mm["width_mm"] - margin_left - margin_mm["right_mm"],
                    height=5.0
                )

            # 제목 판별
            is_head, head_level = is_heading(text, para.style_id)

            if text:
                elem_type = "heading" if is_head else "paragraph"
                elem = DocumentElement(
                    element_id=next_elem_id(),
                    element_type=elem_type,
                    text=text,
                    bbox=bbox,
                    page=current_page_num,
                    level=head_level,
                    style={"style_id": para.style_id},
                )
                elements.append(elem)

                if is_head:
                    headings.append(elem)

                    # 계층 구조 구축
                    hier_section = HierarchicalSection(
                        section_id=next_section_id(),
                        title=text,
                        level=head_level,
                        content=[],
                        tables=[],
                        bbox=bbox,
                        page=current_page_num,
                    )

                    # 상위 레벨 찾기
                    if head_level == 1:
                        root_sections.append(hier_section)
                        current_sections = {1: hier_section, 2: None, 3: None}
                    elif head_level == 2:
                        if current_sections[1]:
                            current_sections[1].children.append(hier_section)
                        else:
                            root_sections.append(hier_section)
                        current_sections[2] = hier_section
                        current_sections[3] = None
                    elif head_level == 3:
                        if current_sections[2]:
                            current_sections[2].children.append(hier_section)
                        elif current_sections[1]:
                            current_sections[1].children.append(hier_section)
                        else:
                            root_sections.append(hier_section)
                        current_sections[3] = hier_section
                else:
                    paragraphs.append(elem)

                    # 현재 섹션에 내용 추가
                    for level in [3, 2, 1]:
                        if current_sections[level]:
                            current_sections[level].content.append(text)
                            break

            # 테이블 처리
            for table in para.tables:
                table_id = next_elem_id()

                # 테이블 바운딩 박스
                size_mm = table.size.to_mm()
                pos_mm = table.position.to_mm()

                if table.position.treat_as_char:
                    table_x = margin_left + pos_mm["horz_offset_mm"]
                    table_y = bbox.y if bbox.is_valid() else margin_top
                else:
                    table_x = margin_left + pos_mm["horz_offset_mm"]
                    table_y = margin_top + pos_mm["vert_offset_mm"]

                table_bbox = BoundingBox(
                    x=table_x,
                    y=table_y,
                    width=size_mm["width_mm"],
                    height=size_mm["height_mm"],
                )

                # 셀 데이터 추출
                grid = [[None for _ in range(table.cols)] for _ in range(table.rows)]
                for cell in table.cells:
                    if 0 <= cell.row < table.rows and 0 <= cell.col < table.cols:
                        grid[cell.row][cell.col] = cell.text

                for r in range(table.rows):
                    for c in range(table.cols):
                        if grid[r][c] is None:
                            grid[r][c] = ""

                # 헤더/데이터 분리
                headers = [grid[0]] if grid else []
                rows = grid[1:] if len(grid) > 1 else []

                # 표 제목 추출
                table_title = extract_table_title(prev_para_text)

                table_struct = TableStructure(
                    table_id=table_id,
                    title=table_title,
                    headers=headers,
                    rows=rows,
                    bbox=table_bbox,
                    page=current_page_num,
                    row_count=table.rows,
                    col_count=table.cols,
                    context=prev_para_text[:200] if prev_para_text else "",
                )
                tables.append(table_struct)

                # 현재 섹션에 표 추가
                for level in [3, 2, 1]:
                    if current_sections[level]:
                        current_sections[level].tables.append(table_struct)
                        break

                # 테이블 요소
                table_elem = DocumentElement(
                    element_id=table_id,
                    element_type="table",
                    text=f"[표 {table.rows}×{table.cols}] {table_title}",
                    bbox=table_bbox,
                    page=current_page_num,
                )
                elements.append(table_elem)

            prev_para_text = text

    return ExtractedDocument(
        title=doc.title,
        source_file=str(doc.folder_path),
        file_type="hwpx",
        pages=pages,
        elements=elements,
        tables=tables,
        headings=headings,
        paragraphs=paragraphs,
        images=[],  # 이미지는 extract_document_with_images에서 추가
        hierarchical_sections=root_sections,
        metadata={
            "version": f"{doc.version.application} {doc.version.app_version}",
        }
    )


# =============================================================================
# HWP 문서 추출
# =============================================================================

def extract_from_hwp(doc) -> ExtractedDocument:
    """
    HWP 문서에서 구조화된 정보 추출

    Args:
        doc: HwpDocument 객체

    Returns:
        ExtractedDocument
    """
    elements = []
    tables = []
    headings = []
    paragraphs = []
    pages = []
    hierarchical_sections = []

    element_counter = 0
    section_counter = 0

    def next_elem_id() -> str:
        nonlocal element_counter
        element_counter += 1
        return f"elem_{element_counter:04d}"

    def next_section_id() -> str:
        nonlocal section_counter
        section_counter += 1
        return f"sec_{section_counter:04d}"

    # 현재 계층 구조 추적
    current_sections = {1: None, 2: None, 3: None}
    root_sections = []

    for section in doc.sections:
        # 페이지 정보
        page_width = section.page_width_mm()
        page_height = section.page_height_mm()
        margin_left = section.margin_left_mm()
        margin_top = section.margin_top_mm()
        margin_right = section.margin_right * HWPUNIT_TO_MM if section.margin_right else 20.0
        margin_bottom = section.margin_bottom * HWPUNIT_TO_MM if section.margin_bottom else 20.0
        
        # 콘텐츠 영역 높이
        content_height = page_height - margin_top - margin_bottom
        
        # 첫 페이지 생성
        current_page_num = len(pages)
        page_info = PageInfo(
            page_num=current_page_num,
            width=page_width,
            height=page_height,
            margin_left=margin_left,
            margin_top=margin_top,
            margin_right=margin_right,
            margin_bottom=margin_bottom,
        )
        pages.append(page_info)
        
        # 현재 Y 위치 추적
        current_y = margin_top
        line_height = 5.0  # 기본 줄 높이
        prev_para_text = ""

        for para in section.paragraphs:
            text = para.plain_text.strip()

            if not text:
                continue

            # 텍스트 높이 추정
            line_count = text.count('\n') + 1
            text_height = line_count * line_height
            
            # 페이지 경계 체크 - 넘어가면 새 페이지로
            if current_y + text_height > page_height - margin_bottom:
                current_page_num = len(pages)
                current_y = margin_top
                
                # 새 페이지 정보 추가
                new_page_info = PageInfo(
                    page_num=current_page_num,
                    width=page_width,
                    height=page_height,
                    margin_left=margin_left,
                    margin_top=margin_top,
                    margin_right=margin_right,
                    margin_bottom=margin_bottom,
                )
                pages.append(new_page_info)

            # 바운딩 박스 (파서에서 계산된 것 또는 새로 계산)
            if para.bbox.is_valid():
                bbox = BoundingBox(
                    x=para.bbox.x,
                    y=current_y,  # 현재 Y 위치 사용
                    width=para.bbox.width,
                    height=para.bbox.height if para.bbox.height > 0 else text_height,
                )
            else:
                bbox = BoundingBox(
                    x=margin_left,
                    y=current_y,
                    width=page_width - margin_left - margin_right,
                    height=text_height,
                )

            # 제목 판별
            is_head, head_level = is_heading(text)

            elem_type = "heading" if is_head else "paragraph"
            elem = DocumentElement(
                element_id=next_elem_id(),
                element_type=elem_type,
                text=text,
                bbox=bbox,
                page=current_page_num,  # 현재 페이지 번호 사용
                level=head_level,
            )
            elements.append(elem)
            
            # Y 위치 업데이트
            current_y += bbox.height + 2.0  # 문단 간격

            if is_head:
                headings.append(elem)

                # 계층 구조 구축
                hier_section = HierarchicalSection(
                    section_id=next_section_id(),
                    title=text,
                    level=head_level,
                    content=[],
                    tables=[],
                    bbox=bbox,
                    page=current_page_num,  # 현재 페이지 번호 사용
                )

                if head_level == 1:
                    root_sections.append(hier_section)
                    current_sections = {1: hier_section, 2: None, 3: None}
                elif head_level == 2:
                    if current_sections[1]:
                        current_sections[1].children.append(hier_section)
                    else:
                        root_sections.append(hier_section)
                    current_sections[2] = hier_section
                    current_sections[3] = None
                elif head_level == 3:
                    if current_sections[2]:
                        current_sections[2].children.append(hier_section)
                    elif current_sections[1]:
                        current_sections[1].children.append(hier_section)
                    else:
                        root_sections.append(hier_section)
                    current_sections[3] = hier_section
            else:
                paragraphs.append(elem)

                for level in [3, 2, 1]:
                    if current_sections[level]:
                        current_sections[level].content.append(text)
                        break

            # 테이블 처리
            for table in para.tables:
                table_id = next_elem_id()

                table_bbox = BoundingBox(
                    x=margin_left,
                    y=bbox.y + bbox.height,
                    width=page_info.width - margin_left - page_info.margin_right,
                    height=table.rows * 5.0,
                )

                grid = [[None for _ in range(table.cols)] for _ in range(table.rows)]
                for cell in table.cells:
                    if 0 <= cell.row < table.rows and 0 <= cell.col < table.cols:
                        grid[cell.row][cell.col] = cell.text

                for r in range(table.rows):
                    for c in range(table.cols):
                        if grid[r][c] is None:
                            grid[r][c] = ""

                headers = [grid[0]] if grid else []
                rows = grid[1:] if len(grid) > 1 else []

                table_title = extract_table_title(prev_para_text)

                table_struct = TableStructure(
                    table_id=table_id,
                    title=table_title,
                    headers=headers,
                    rows=rows,
                    bbox=table_bbox,
                    page=section.index,
                    row_count=table.rows,
                    col_count=table.cols,
                )
                tables.append(table_struct)

                for level in [3, 2, 1]:
                    if current_sections[level]:
                        current_sections[level].tables.append(table_struct)
                        break

            prev_para_text = text

    return ExtractedDocument(
        title=doc.title,
        source_file=str(doc.file_path),
        file_type="hwp",
        pages=pages,
        elements=elements,
        tables=tables,
        headings=headings,
        paragraphs=paragraphs,
        images=[],  # 이미지는 extract_document_with_images에서 추가
        hierarchical_sections=root_sections,
        metadata={
            "version": doc.header.version,
        }
    )


def extract_images_to_elements(
    file_path: str | Path,
    file_type: str,
) -> list[ImageElement]:
    """
    파일에서 이미지를 추출하여 ImageElement 목록으로 반환
    
    Args:
        file_path: HWP 또는 HWPX 파일 경로
        file_type: "hwp" 또는 "hwpx"
        
    Returns:
        ImageElement 목록
    """
    if not HAS_IMAGE_EXTRACTOR:
        return []
    
    image_elements = []
    
    try:
        if file_type == "hwp":
            images = extract_images_from_hwp(file_path)
        elif file_type == "hwpx":
            images = extract_images_from_hwpx(file_path)
        else:
            return []
        
        for img in images:
            img_elem = ImageElement(
                image_id=img.bin_id,
                filename=img.filename,
                format=img.format,
                bbox=BoundingBox(
                    x=img.x,
                    y=img.y,
                    width=img.width,
                    height=img.height,
                ),
                page=img.page,
                pixel_width=img.pixel_width,
                pixel_height=img.pixel_height,
                file_size=len(img.data),
            )
            image_elements.append(img_elem)
            
    except Exception as e:
        print(f"이미지 추출 오류: {e}")
    
    return image_elements


def extract_document_with_images(
    doc,
    extract_images: bool = True,
    save_images_dir: Optional[Path] = None,
) -> ExtractedDocument:
    """
    문서에서 구조화된 정보와 이미지를 함께 추출
    
    Args:
        doc: HwpxDocument 또는 HwpDocument 객체
        extract_images: 이미지 추출 여부
        save_images_dir: 이미지 저장 디렉토리 (None이면 저장하지 않음)
        
    Returns:
        ExtractedDocument (이미지 포함)
    """
    # 기본 문서 추출
    extracted = extract_document(doc)
    
    # 이미지 추출
    if extract_images and HAS_IMAGE_EXTRACTOR:
        file_path = Path(extracted.source_file)
        
        try:
            if extracted.file_type == "hwp":
                images = extract_images_from_hwp(file_path)
            elif extracted.file_type == "hwpx":
                images = extract_images_from_hwpx(file_path)
            else:
                images = []
            
            # ImageElement로 변환
            for img in images:
                saved_path = ""
                
                # 이미지 저장
                if save_images_dir:
                    saved_path = str(img.save(save_images_dir, convert_vector=True))
                
                img_elem = ImageElement(
                    image_id=img.bin_id,
                    filename=img.filename,
                    format=img.format,
                    bbox=BoundingBox(
                        x=img.x,
                        y=img.y,
                        width=img.width,
                        height=img.height,
                    ),
                    page=img.page,
                    pixel_width=img.pixel_width,
                    pixel_height=img.pixel_height,
                    file_size=len(img.data),
                    saved_path=saved_path,
                )
                extracted.images.append(img_elem)
                
        except Exception as e:
            print(f"이미지 추출 오류: {e}")
    
    return extracted


# =============================================================================
# 통합 추출 함수
# =============================================================================

def extract_document(doc) -> ExtractedDocument:
    """
    문서에서 구조화된 정보 추출

    HWPX와 HWP 문서를 자동으로 감지하여 적절한 함수를 호출합니다.

    Args:
        doc: HwpxDocument 또는 HwpDocument 객체

    Returns:
        ExtractedDocument
    """
    # HWPX 문서 감지
    if hasattr(doc, 'sections') and hasattr(doc, 'version'):
        if doc.sections and hasattr(doc.sections[0], 'page_props'):
            return extract_from_hwpx(doc)

    # HWP 문서 감지
    if hasattr(doc, 'header') and hasattr(doc.header, 'is_compressed'):
        return extract_from_hwp(doc)

    raise ValueError("지원하지 않는 문서 형식입니다.")


# =============================================================================
# 시각화 함수
# =============================================================================

def visualize_elements(
    extracted: ExtractedDocument,
    output_path: str | Path,
    page_num: int = 0,
    show_bbox: bool = True,
    show_text: bool = True,
    scale: float = 3.0,
    font_size: int = 8,
) -> Path:
    """
    추출된 문서 요소 시각화 (원본 문서와 유사한 레이아웃)

    Args:
        extracted: ExtractedDocument 객체
        output_path: 출력 이미지 경로
        page_num: 페이지 번호
        show_bbox: 바운딩 박스 표시 여부
        show_text: 텍스트 표시 여부
        scale: 확대 비율
        font_size: 폰트 크기

    Returns:
        저장된 이미지 경로
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        raise ImportError("Pillow 라이브러리가 필요합니다. pip install Pillow")

    if page_num >= len(extracted.pages):
        raise ValueError(f"페이지 {page_num}이 존재하지 않습니다.")

    page = extracted.pages[page_num]
    page_elements = [e for e in extracted.elements if e.page == page_num]

    # 이미지 크기
    img_width = int(page.width * scale)
    img_height = int(page.height * scale)

    img = Image.new('RGB', (img_width, img_height), 'white')
    draw = ImageDraw.Draw(img)

    # 폰트 로드
    def load_font(size):
        try:
            return ImageFont.truetype("/System/Library/Fonts/AppleSDGothicNeo.ttc", size)
        except:
            try:
                return ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", size)
            except:
                return ImageFont.load_default()

    font = load_font(font_size)
    small_font = load_font(max(6, font_size - 2))
    title_font = load_font(font_size + 2)

    # 클래스 코드 매핑
    CLASS_CODES = {
        "heading": "h",
        "paragraph": "p",
        "table": "T",
        "table_cell": "c",
        "image": "i",
    }

    # 색상 (반투명 효과를 위한 연한 색상)
    colors = {
        "heading": {"outline": "#C2185B", "fill": "#FCE4EC", "text": "#880E4F"},
        "paragraph": {"outline": "#1976D2", "fill": "#E3F2FD", "text": "#0D47A1"},
        "table": {"outline": "#388E3C", "fill": "#E8F5E9", "text": "#1B5E20"},
        "table_cell": {"outline": "#F57C00", "fill": "#FFF3E0", "text": "#E65100"},
        "image": {"outline": "#7B1FA2", "fill": "#F3E5F5", "text": "#4A148C"},
    }

    # 페이지 테두리
    draw.rectangle([(0, 0), (img_width - 1, img_height - 1)], outline='#333333', width=2)

    # 여백 영역 (점선 효과)
    margin_rect = [
        int(page.margin_left * scale),
        int(page.margin_top * scale),
        int((page.width - page.margin_right) * scale),
        int((page.height - page.margin_bottom) * scale),
    ]
    draw.rectangle(margin_rect, outline='#CCCCCC', width=1)

    # 요소 그리기 (레이아웃 박스 안에 텍스트 포함)
    for elem in page_elements:
        if not elem.bbox.is_valid():
            continue

        color = colors.get(elem.element_type, colors["paragraph"])
        class_code = CLASS_CODES.get(elem.element_type, "?")

        x1 = int(elem.bbox.x * scale)
        y1 = int(elem.bbox.y * scale)
        x2 = int((elem.bbox.x + elem.bbox.width) * scale)
        y2 = int((elem.bbox.y + elem.bbox.height) * scale)

        # 최소 박스 크기 보장
        if x2 - x1 < 10:
            x2 = x1 + 10
        if y2 - y1 < 10:
            y2 = y1 + 10

        box_width = x2 - x1
        box_height = y2 - y1

        if show_bbox:
            # 배경색으로 채운 박스
            draw.rectangle([(x1, y1), (x2, y2)], fill=color["fill"], outline=color["outline"], width=1)

        # 텍스트 표시
        if show_text:
            text = elem.text.strip() if elem.text else ""

            # 클래스 코드 표시 (좌상단 작은 태그)
            code_text = f"[{class_code}]"
            try:
                draw.text((x1 + 1, y1), code_text, fill=color["text"], font=small_font)
            except:
                pass

            # 텍스트 표시 (박스 크기에 맞게 줄바꿈)
            if text:
                # 사용 가능한 텍스트 영역
                text_x = x1 + 20  # 클래스 코드 뒤
                text_y = y1 + 1
                available_width = max(50, box_width - 22)  # 최소 50px 확보
                available_height = max(10, box_height - 2)

                # 한 줄에 들어갈 수 있는 대략적인 문자 수 (한글 기준)
                char_width = font_size * 0.55  # 대략적인 문자 너비
                chars_per_line = max(5, int(available_width / char_width))

                # 여러 줄로 텍스트 분할
                lines = []
                remaining = text
                line_height = font_size + 1
                max_lines = max(1, int(available_height / line_height))

                for i in range(max_lines):
                    if not remaining:
                        break
                    if len(remaining) <= chars_per_line:
                        lines.append(remaining)
                        remaining = ""
                    else:
                        lines.append(remaining[:chars_per_line])
                        remaining = remaining[chars_per_line:]

                # 남은 텍스트가 있으면 마지막 줄에 ... 추가
                if remaining and lines:
                    last_line = lines[-1]
                    if len(last_line) > 3:
                        lines[-1] = last_line[:-3] + "..."
                    else:
                        lines[-1] = last_line + ".."

                # 텍스트 그리기
                current_y = text_y
                for line in lines:
                    try:
                        draw.text((text_x, current_y), line, fill='black', font=small_font)
                    except:
                        pass
                    current_y += line_height

    # 제목 (상단)
    title = f"{extracted.title} - Page {page_num + 1}/{len(extracted.pages)}"
    draw.rectangle([(0, 0), (img_width, 18)], fill='#F5F5F5')
    draw.text((5, 2), title, fill='#333333', font=title_font)

    # 범례 (우하단)
    legend_x = img_width - 120
    legend_y = img_height - 65
    draw.rectangle([(legend_x - 5, legend_y - 5), (img_width - 5, img_height - 5)], fill='#FAFAFA', outline='#CCCCCC')
    draw.text((legend_x, legend_y), "Classes:", fill='#333333', font=small_font)
    legend_y += 12
    legend_items = [
        ("h", "heading", colors["heading"]),
        ("p", "paragraph", colors["paragraph"]),
        ("T", "table", colors["table"]),
        ("c", "cell", colors["table_cell"]),
    ]
    for code, name, color in legend_items:
        draw.rectangle([(legend_x, legend_y), (legend_x + 10, legend_y + 8)], fill=color["fill"], outline=color["outline"])
        draw.text((legend_x + 13, legend_y - 1), f"{code}={name}", fill='#333333', font=small_font)
        legend_y += 11

    output_path = Path(output_path)
    img.save(output_path)
    print(f"시각화 저장: {output_path}")

    return output_path


def visualize_all_pages(
    extracted: ExtractedDocument,
    output_dir: str | Path,
    show_bbox: bool = True,
    show_text: bool = True,
    scale: float = 3.0,
    font_size: int = 10,
) -> list[Path]:
    """
    모든 페이지를 시각화합니다.

    Args:
        extracted: ExtractedDocument 객체
        output_dir: 출력 디렉토리
        show_bbox: 바운딩 박스 표시 여부
        show_text: 텍스트 표시 여부
        scale: 확대 비율
        font_size: 폰트 크기

    Returns:
        list[Path]: 생성된 이미지 파일 경로 리스트
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    saved_files = []

    for page_num in range(len(extracted.pages)):
        output_path = output_dir / f"{extracted.title}_page_{page_num + 1:03d}.png"
        try:
            visualize_elements(
                extracted,
                output_path,
                page_num=page_num,
                show_bbox=show_bbox,
                show_text=show_text,
                scale=scale,
                font_size=font_size,
            )
            saved_files.append(output_path)
        except Exception as e:
            print(f"페이지 {page_num + 1} 시각화 실패: {e}")

    print(f"총 {len(saved_files)}개 페이지 시각화 완료")
    return saved_files


def visualize_to_pdf(
    extracted: ExtractedDocument,
    output_path: str | Path,
    show_bbox: bool = True,
    show_text: bool = True,
    scale: float = 3.0,
    font_size: int = 10,
) -> Path:
    """
    모든 페이지를 하나의 PDF로 시각화합니다 (레이아웃 중심).

    Args:
        extracted: ExtractedDocument 객체
        output_path: 출력 PDF 경로
        show_bbox: 바운딩 박스 표시 여부
        show_text: 텍스트 표시 여부
        scale: 확대 비율
        font_size: 폰트 크기

    Returns:
        Path: 생성된 PDF 파일 경로
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        raise ImportError("Pillow 라이브러리가 필요합니다. pip install Pillow")

    # 클래스 코드 매핑
    CLASS_CODES = {
        "heading": "h",
        "paragraph": "p",
        "table": "T",
        "table_cell": "c",
        "image": "i",
    }

    images = []

    for page_num in range(len(extracted.pages)):
        page = extracted.pages[page_num]
        page_elements = [e for e in extracted.elements if e.page == page_num]

        # 이미지 크기
        img_width = int(page.width * scale)
        img_height = int(page.height * scale)

        img = Image.new('RGB', (img_width, img_height), 'white')
        draw = ImageDraw.Draw(img)

        # 폰트
        try:
            font = ImageFont.truetype("/System/Library/Fonts/AppleSDGothicNeo.ttc", font_size)
            small_font = ImageFont.truetype("/System/Library/Fonts/AppleSDGothicNeo.ttc", font_size - 2)
        except:
            try:
                font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", font_size)
                small_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", font_size - 2)
            except:
                font = ImageFont.load_default()
                small_font = font

        # 색상
        colors = {
            "heading": {"outline": "#E91E63", "fill": "#FCE4EC", "text": "#880E4F"},
            "paragraph": {"outline": "#2196F3", "fill": "#E3F2FD", "text": "#0D47A1"},
            "table": {"outline": "#4CAF50", "fill": "#E8F5E9", "text": "#1B5E20"},
            "table_cell": {"outline": "#FF9800", "fill": "#FFF3E0", "text": "#E65100"},
            "image": {"outline": "#9C27B0", "fill": "#F3E5F5", "text": "#4A148C"},
        }

        # 페이지 테두리
        draw.rectangle([(0, 0), (img_width - 1, img_height - 1)], outline='black', width=2)

        # 여백 영역
        margin_rect = [
            int(page.margin_left * scale),
            int(page.margin_top * scale),
            int((page.width - page.margin_right) * scale),
            int((page.height - page.margin_bottom) * scale),
        ]
        draw.rectangle(margin_rect, outline='lightgray', width=1)

        # 요소 그리기
        for elem in page_elements:
            if not elem.bbox.is_valid():
                continue

            color = colors.get(elem.element_type, colors["paragraph"])
            class_code = CLASS_CODES.get(elem.element_type, "?")

            x1 = int(elem.bbox.x * scale)
            y1 = int(elem.bbox.y * scale)
            x2 = int((elem.bbox.x + elem.bbox.width) * scale)
            y2 = int((elem.bbox.y + elem.bbox.height) * scale)

            box_width = x2 - x1

            if show_bbox:
                draw.rectangle([(x1, y1), (x2, y2)], fill=color["fill"], outline=color["outline"], width=1)

            if show_text:
                text = elem.text.strip() if elem.text else ""
                code_text = f"[{class_code}]"
                try:
                    draw.text((x1 + 2, y1 + 1), code_text, fill=color["text"], font=small_font)
                except:
                    pass

                if text:
                    max_chars = max(1, (box_width - 25) // (font_size // 2))
                    if len(text) > max_chars:
                        text = text[:max_chars - 2] + ".."
                    try:
                        draw.text((x1 + 22, y1 + 1), text, fill='black', font=small_font)
                    except:
                        pass

        # 제목
        title = f"{extracted.title} - Page {page_num + 1}/{len(extracted.pages)}"
        draw.text((10, 5), title, fill='black', font=font)

        # 범례
        legend_y = img_height - 70
        draw.text((10, legend_y), "Classes:", fill='black', font=font)
        legend_y += 14
        legend_items = [
            ("h", "heading", colors["heading"]),
            ("p", "paragraph", colors["paragraph"]),
            ("T", "table", colors["table"]),
            ("c", "cell", colors["table_cell"]),
        ]
        for code, name, color in legend_items:
            draw.rectangle([(10, legend_y), (22, legend_y + 10)], fill=color["fill"], outline=color["outline"])
            draw.text((25, legend_y - 1), f"{code}={name}", fill='black', font=small_font)
            legend_y += 12

        images.append(img)

    # PDF로 저장
    output_path = Path(output_path)
    if images:
        images[0].save(
            output_path,
            "PDF",
            save_all=True,
            append_images=images[1:] if len(images) > 1 else [],
            resolution=100.0,
        )
        print(f"PDF 시각화 저장: {output_path} ({len(images)}페이지)")

    return output_path


def create_visualization_report(
    extracted: ExtractedDocument,
    output_dir: str | Path,
) -> list[Path]:
    """
    전체 문서 시각화 리포트 생성

    Args:
        extracted: ExtractedDocument 객체
        output_dir: 출력 디렉토리

    Returns:
        생성된 파일 경로 리스트
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    saved_files = []

    # 각 페이지 시각화
    for page_num in range(len(extracted.pages)):
        img_path = output_dir / f"{extracted.title}_page_{page_num + 1:03d}.png"
        try:
            visualize_elements(extracted, img_path, page_num=page_num)
            saved_files.append(img_path)
        except Exception as e:
            print(f"페이지 {page_num + 1} 시각화 실패: {e}")

    # JSON 저장
    json_path = output_dir / f"{extracted.title}_extracted.json"
    with open(json_path, "w", encoding="utf-8") as f:
        f.write(extracted.to_json())
    saved_files.append(json_path)
    print(f"JSON 저장: {json_path}")

    # 구조화된 텍스트 저장
    txt_path = output_dir / f"{extracted.title}_structured.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(extracted.to_structured_text())
    saved_files.append(txt_path)
    print(f"구조화된 텍스트 저장: {txt_path}")

    # RAG 청크 저장
    chunks_path = output_dir / f"{extracted.title}_chunks.json"
    chunks = extracted.to_rag_chunks()
    with open(chunks_path, "w", encoding="utf-8") as f:
        json.dump(chunks, f, ensure_ascii=False, indent=2)
    saved_files.append(chunks_path)
    print(f"RAG 청크 저장: {chunks_path}")

    # 표 목록 저장
    if extracted.tables:
        tables_path = output_dir / f"{extracted.title}_tables.md"
        with open(tables_path, "w", encoding="utf-8") as f:
            f.write(f"# {extracted.title} - 표 목록\n\n")
            f.write(f"**총 {len(extracted.tables)}개 표** | **페이지 수: {len(extracted.pages)}**\n\n")
            for i, table in enumerate(extracted.tables):
                page_num = table.page + 1  # 1-indexed for display
                f.write(f"## 표 {i + 1} (페이지 {page_num})\n\n")
                f.write(table.to_markdown())
                f.write("\n\n")
        saved_files.append(tables_path)
        print(f"표 목록 저장: {tables_path}")

    # 이미지 목록 저장 (OCR 연동용)
    if extracted.images:
        # 이미지 메타데이터 JSON (외부 OCR 연동용)
        images_json_path = output_dir / f"{extracted.title}_images.json"
        images_data = {
            "document_title": extracted.title,
            "image_count": len(extracted.images),
            "images": [img.to_ocr_dict() for img in extracted.images],
        }
        with open(images_json_path, "w", encoding="utf-8") as f:
            json.dump(images_data, f, ensure_ascii=False, indent=2)
        saved_files.append(images_json_path)
        print(f"이미지 메타데이터 저장: {images_json_path}")
        
        # 이미지 목록 마크다운
        images_md_path = output_dir / f"{extracted.title}_images.md"
        with open(images_md_path, "w", encoding="utf-8") as f:
            f.write(f"# {extracted.title} - 이미지 목록\n\n")
            f.write(f"**총 {len(extracted.images)}개 이미지** | **페이지 수: {len(extracted.pages)}**\n\n")
            f.write("| # | 파일명 | 형식 | 크기 | 해상도 | 위치 (mm) | 페이지 |\n")
            f.write("|---|--------|------|------|--------|-----------|--------|\n")
            for i, img in enumerate(extracted.images, 1):
                size_str = f"{img.file_size:,} B"
                res_str = f"{img.pixel_width}×{img.pixel_height}" if img.pixel_width else "-"
                pos_str = f"({img.bbox.x:.1f}, {img.bbox.y:.1f})" if img.bbox.width > 0 else "-"
                page_str = str(img.page + 1) if img.bbox.width > 0 else "-"
                f.write(f"| {i} | {img.filename} | {img.format.upper()} | {size_str} | {res_str} | {pos_str} | {page_str} |\n")
            f.write("\n")
            
            # 상세 정보
            f.write("## 상세 정보\n\n")
            for i, img in enumerate(extracted.images, 1):
                f.write(f"### {i}. {img.filename}\n\n")
                f.write(f"- **이미지 ID**: `{img.image_id}`\n")
                f.write(f"- **형식**: {img.format.upper()}\n")
                f.write(f"- **파일 크기**: {img.file_size:,} bytes\n")
                if img.pixel_width and img.pixel_height:
                    f.write(f"- **해상도**: {img.pixel_width}×{img.pixel_height} px\n")
                if img.bbox.width > 0:
                    f.write(f"- **문서 내 위치**: ({img.bbox.x:.2f}, {img.bbox.y:.2f}) mm\n")
                    f.write(f"- **문서 내 크기**: {img.bbox.width:.2f}×{img.bbox.height:.2f} mm\n")
                    f.write(f"- **페이지**: {img.page + 1}\n")
                if img.saved_path:
                    f.write(f"- **저장 경로**: `{img.saved_path}`\n")
                f.write("\n")
                
        saved_files.append(images_md_path)
        print(f"이미지 목록 저장: {images_md_path}")

    return saved_files


# =============================================================================
# 메인
# =============================================================================

if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).parent))

    from hwpx_parser import parse_hwpx
    from hwp_parser import parse_hwp

    # 테스트 파일
    data_dir = Path(__file__).parent.parent / "data" / "docs"
    output_dir = Path(__file__).parent / "results"

    hwpx_file = data_dir / "은행권 광고심의 결과 보고서(양식)vF (1).hwpx"
    hwp_file = data_dir / "2. [농협] 광고안(B).hwp"

    print("=" * 70)
    print("Document Extractor 테스트")
    print("=" * 70)

    # HWPX 테스트
    if hwpx_file.exists():
        print(f"\nHWPX 파일 처리: {hwpx_file.name}")
        doc = parse_hwpx(hwpx_file)
        
        # 이미지 저장 디렉토리
        hwpx_img_dir = output_dir / "hwpx_extracted" / "images"
        
        # 이미지 포함 추출
        extracted = extract_document_with_images(
            doc, 
            extract_images=True, 
            save_images_dir=hwpx_img_dir
        )

        print(f"  - 요소 수: {len(extracted.elements)}")
        print(f"  - 제목 수: {len(extracted.headings)}")
        print(f"  - 표 수: {len(extracted.tables)}")
        print(f"  - 이미지 수: {len(extracted.images)}")
        print(f"  - 계층 섹션 수: {len(extracted.hierarchical_sections)}")

        # 시각화
        create_visualization_report(extracted, output_dir / "hwpx_extracted")

    # HWP 테스트
    if hwp_file.exists():
        print(f"\nHWP 파일 처리: {hwp_file.name}")
        doc = parse_hwp(hwp_file)
        
        # 이미지 저장 디렉토리
        hwp_img_dir = output_dir / "hwp_extracted" / "images"
        
        # 이미지 포함 추출
        extracted = extract_document_with_images(
            doc, 
            extract_images=True, 
            save_images_dir=hwp_img_dir
        )

        print(f"  - 요소 수: {len(extracted.elements)}")
        print(f"  - 제목 수: {len(extracted.headings)}")
        print(f"  - 표 수: {len(extracted.tables)}")
        print(f"  - 이미지 수: {len(extracted.images)}")

        # 시각화
        create_visualization_report(extracted, output_dir / "hwp_extracted")

    print("\n테스트 완료!")
