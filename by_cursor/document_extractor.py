"""
Document Extractor - HWPX/HWP 문서에서 구조화된 정보 추출

이 모듈은 파싱된 문서에서 LLM/RAG에 적합한 형태로 정보를 추출합니다.

주요 기능:
1. 정확한 바운딩 박스 좌표 추출 (페이지 내 절대 좌표)
2. 표의 구조화 (제목/헤더/내용 분리)
3. 문서 요소의 계층적 구조화
4. 시각화를 위한 뷰어 함수

사용 예시:
    from document_extractor import extract_document_elements, visualize_elements
    from hwpx_parser_cursor import parse_hwpx
    
    # 파싱
    doc = parse_hwpx("document.hwpx")
    
    # 구조화된 정보 추출
    elements = extract_document_elements(doc)
    
    # 시각화
    visualize_elements(elements, "output.png")
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any, Literal, Optional
from pathlib import Path
import json

# 이미지 추출 모듈 (선택적)
try:
    from image_extractor import (
        extract_images_from_hwp, 
        extract_images_from_hwpx,
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
    """
    바운딩 박스 (절대 좌표)
    
    Attributes:
        x: X 좌표 (mm, 페이지 왼쪽 상단 기준)
        y: Y 좌표 (mm, 페이지 왼쪽 상단 기준)
        width: 너비 (mm)
        height: 높이 (mm)
    """
    x: float
    y: float
    width: float
    height: float
    
    def to_dict(self) -> dict:
        return {
            "x": round(self.x, 2),
            "y": round(self.y, 2),
            "width": round(self.width, 2),
            "height": round(self.height, 2),
            "x2": round(self.x + self.width, 2),
            "y2": round(self.y + self.height, 2),
        }
    
    def __repr__(self):
        return f"BBox({self.x:.1f}, {self.y:.1f}, {self.width:.1f}×{self.height:.1f})"


@dataclass
class DocumentElement:
    """
    문서 요소 (텍스트, 표, 이미지 등)
    
    Attributes:
        element_type: 요소 유형 (text, table, table_cell, image, heading)
        text: 텍스트 내용
        bbox: 바운딩 박스 (절대 좌표)
        page: 페이지 번호 (0부터)
        level: 요소 레벨 (heading의 경우 1, 2, 3 등)
        parent_id: 부모 요소 ID (테이블 셀의 경우 테이블 ID)
        children: 자식 요소들
        style: 스타일 정보
        metadata: 추가 메타데이터
    """
    element_id: str
    element_type: Literal["text", "heading", "table", "table_cell", "image", "paragraph"]
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
            "bbox": self.bbox.to_dict(),
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
    표의 구조화된 정보
    
    Attributes:
        table_id: 표 ID
        title: 표 제목 (표 위의 텍스트에서 추출)
        headers: 헤더 행 (첫 번째 행 또는 병합된 헤더)
        rows: 데이터 행들
        bbox: 표 전체 바운딩 박스
    """
    table_id: str
    title: str
    headers: list[list[str]]
    rows: list[list[str]]
    bbox: BoundingBox
    page: int = 0
    metadata: dict = field(default_factory=dict)
    
    def to_dict(self) -> dict:
        return {
            "id": self.table_id,
            "title": self.title,
            "headers": self.headers,
            "rows": self.rows,
            "bbox": self.bbox.to_dict(),
            "page": self.page,
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
        
        # 열 개수 맞추기
        max_cols = max(len(row) for row in all_rows) if all_rows else 0
        
        for i, row in enumerate(all_rows):
            # 열 개수 맞추기
            padded_row = row + [""] * (max_cols - len(row))
            lines.append("| " + " | ".join(cell.replace("|", "\\|").replace("\n", " ") for cell in padded_row) + " |")
            
            # 헤더 구분선
            if i == len(self.headers) - 1 and self.headers:
                lines.append("|" + "|".join(["---"] * max_cols) + "|")
        
        return "\n".join(lines)
    
    def to_structured_text(self) -> str:
        """LLM에 적합한 구조화된 텍스트로 변환"""
        lines = []
        
        if self.title:
            lines.append(f"[표 제목] {self.title}")
        
        if self.headers:
            header_text = " | ".join(self.headers[0]) if self.headers[0] else ""
            lines.append(f"[표 헤더] {header_text}")
        
        for i, row in enumerate(self.rows):
            row_text = " | ".join(row)
            lines.append(f"[행 {i+1}] {row_text}")
        
        return "\n".join(lines)


@dataclass
class PageInfo:
    """페이지 정보"""
    page_num: int
    width: float  # mm
    height: float  # mm
    margin_top: float = 0
    margin_bottom: float = 0
    margin_left: float = 0
    margin_right: float = 0
    
    def to_dict(self) -> dict:
        return {
            "page_num": self.page_num,
            "width": self.width,
            "height": self.height,
            "margins": {
                "top": self.margin_top,
                "bottom": self.margin_bottom,
                "left": self.margin_left,
                "right": self.margin_right,
            }
        }


@dataclass
class ImageElement:
    """
    이미지 요소 (외부 OCR 연동용)
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
    ocr_text: str = ""
    
    def to_dict(self) -> dict:
        return {
            "image_id": self.image_id,
            "filename": self.filename,
            "format": self.format,
            "bbox": self.bbox.to_dict(),
            "page": self.page,
            "pixel_width": self.pixel_width,
            "pixel_height": self.pixel_height,
            "file_size": self.file_size,
            "saved_path": self.saved_path,
            "ocr_text": self.ocr_text,
        }
    
    def to_ocr_dict(self) -> dict:
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
    추출된 문서 정보
    
    LLM/RAG에 사용하기 적합한 구조화된 문서 정보
    """
    title: str
    source_file: str
    file_type: str  # "hwpx" or "hwp"
    pages: list[PageInfo]
    elements: list[DocumentElement]
    tables: list[TableStructure]
    headings: list[DocumentElement]
    paragraphs: list[DocumentElement]
    images: list[ImageElement] = field(default_factory=list)
    metadata: dict = field(default_factory=dict)
    
    def to_dict(self) -> dict:
        return {
            "title": self.title,
            "source_file": self.source_file,
            "file_type": self.file_type,
            "page_count": len(self.pages),
            "element_count": len(self.elements),
            "table_count": len(self.tables),
            "image_count": len(self.images),
            "pages": [p.to_dict() for p in self.pages],
            "elements": [e.to_dict() for e in self.elements],
            "tables": [t.to_dict() for t in self.tables],
            "headings": [h.to_dict() for h in self.headings],
            "images": [i.to_dict() for i in self.images],
            "metadata": self.metadata,
        }
    
    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)
    
    def to_structured_text(self) -> str:
        """LLM에 적합한 구조화된 텍스트로 변환"""
        lines = [f"# {self.title}", ""]
        
        # 메타 정보
        lines.append(f"[문서 유형] {self.file_type.upper()}")
        lines.append(f"[페이지 수] {len(self.pages)}")
        lines.append("")
        
        # 제목들 먼저
        if self.headings:
            lines.append("## 문서 구조")
            for h in self.headings:
                indent = "  " * (h.level - 1)
                lines.append(f"{indent}- {h.text}")
            lines.append("")
        
        # 본문 내용
        lines.append("## 본문 내용")
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
            elif elem.element_type == "paragraph":
                if elem.text.strip():
                    lines.append(elem.text.strip())
                    lines.append("")
        
        # 표들
        if self.tables:
            lines.append("## 표 목록")
            lines.append("")
            for table in self.tables:
                lines.append(table.to_structured_text())
                lines.append("")
        
        return "\n".join(lines)
    
    def get_full_text(self) -> str:
        """전체 텍스트 추출"""
        texts = []
        for elem in self.elements:
            if elem.text.strip():
                texts.append(elem.text.strip())
        return "\n".join(texts)


# =============================================================================
# HWPX 문서 추출 함수
# =============================================================================

# HWPUNIT to mm 변환 상수
HWPUNIT_TO_MM = 25.4 / 7200


def extract_from_hwpx(doc) -> ExtractedDocument:
    """
    HWPX 문서에서 구조화된 정보 추출
    
    Args:
        doc: HwpxDocument 객체 (hwpx_parser_cursor에서 파싱된 문서)
    
    Returns:
        ExtractedDocument: 구조화된 문서 정보
    """
    elements = []
    tables = []
    headings = []
    paragraphs = []
    pages = []
    
    element_counter = 0
    
    def next_id() -> str:
        nonlocal element_counter
        element_counter += 1
        return f"elem_{element_counter:04d}"
    
    for section in doc.sections:
        # 페이지 정보
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
        
        # 현재 Y 위치 추적 (절대 좌표 계산용)
        # vert_pos는 문단 내의 상대 좌표이므로, 문단을 순차적으로 쌓아야 함
        current_y = margin_mm["top_mm"]
        prev_para_text = ""  # 표 제목 추출용
        default_line_height = 5.0  # 기본 줄 높이 (mm)
        
        for para in section.paragraphs:
            text = para.full_text.strip()
            
            # 라인 세그먼트에서 바운딩 박스 계산
            # 주의: lineseg의 vert_pos는 문단 내 상대 좌표임 (페이지 기준 X)
            if para.line_segments:
                first_seg = para.line_segments[0]
                last_seg = para.line_segments[-1]
                
                # X 좌표: 여백 + 문단 내 수평 위치
                x = margin_mm["left_mm"] + (first_seg.horz_pos * HWPUNIT_TO_MM)
                
                # Y 좌표: 누적된 current_y 사용 (vert_pos는 문단 내 상대 좌표이므로 무시)
                y = current_y
                
                # 너비: 가장 넓은 라인 세그먼트 기준
                width = max(seg.horz_size * HWPUNIT_TO_MM for seg in para.line_segments)
                
                # 높이: 마지막 라인의 끝 - 첫 라인의 시작
                # vert_pos는 상대 좌표이므로 그대로 높이 계산에 사용
                para_height = (last_seg.vert_pos + last_seg.vert_size) * HWPUNIT_TO_MM
                height = max(para_height, default_line_height)
                
                bbox = BoundingBox(x=x, y=y, width=width, height=height)
            else:
                # 라인 세그먼트가 없으면 기본값
                bbox = BoundingBox(
                    x=margin_mm["left_mm"],
                    y=current_y,
                    width=page_mm["width_mm"] - margin_mm["left_mm"] - margin_mm["right_mm"],
                    height=default_line_height
                )
            
            # 제목 판별 (스타일 ID 또는 텍스트 패턴 기반)
            is_heading = False
            heading_level = 0
            
            # 스타일 ID로 제목 판별
            if para.style_id:
                try:
                    style_num = int(para.style_id)
                    if 1 <= style_num <= 6:
                        is_heading = True
                        heading_level = style_num
                except:
                    pass
            
            # 텍스트 패턴으로 제목 판별 (가. 나. 다. / 1. 2. 3. 등)
            if text and not is_heading:
                import re
                # 한글 가나다 패턴
                if re.match(r'^[가-힣]\.\s', text):
                    is_heading = True
                    heading_level = 2
                # 숫자 패턴
                elif re.match(r'^\d+\.\s', text):
                    is_heading = True
                    heading_level = 2
                # 짧고 굵은 텍스트 (제목일 가능성)
                elif len(text) < 50 and text and not text.endswith('.'):
                    # 글자 크기로 판별 가능하면 추가
                    pass
            
            # 요소 생성 (테이블이 있는 경우 텍스트는 셀 내용이므로 건너뜀)
            has_table = len(para.tables) > 0
            
            if text and not has_table:
                elem_type = "heading" if is_heading else "paragraph"
                elem = DocumentElement(
                    element_id=next_id(),
                    element_type=elem_type,
                    text=text,
                    bbox=bbox,
                    page=section.index,
                    level=heading_level,
                    style={
                        "para_pr_id": para.para_pr_id,
                        "style_id": para.style_id,
                    },
                    metadata={
                        "line_count": len(para.line_segments),
                    }
                )
                elements.append(elem)
                
                if is_heading:
                    headings.append(elem)
                else:
                    paragraphs.append(elem)
                
                # 텍스트가 있는 경우 current_y 업데이트
                current_y = bbox.y + bbox.height + 1.0  # 문단 간격 1mm
            
            # 테이블 처리
            for table in para.tables:
                table_id = next_id()
                
                # 테이블 바운딩 박스
                table_size = table.size.to_mm()
                table_pos = table.position.to_mm()
                
                # 테이블 절대 좌표 계산
                # treat_as_char=True: 텍스트 흐름에 따라 배치
                # treat_as_char=False: 페이지 기준 절대 위치
                table_x = margin_mm["left_mm"] + table_pos["horz_offset_mm"]
                
                # 테이블 Y 좌표는 현재 누적 Y 위치 기준
                # vert_offset은 페이지 상단이 아닌 현재 위치 기준 오프셋
                table_y = current_y + table_pos["vert_offset_mm"]
                
                table_bbox = BoundingBox(
                    x=table_x,
                    y=table_y,
                    width=table_size["width_mm"],
                    height=table_size["height_mm"],
                )
                
                # 테이블 셀 데이터 추출
                grid = [[None for _ in range(table.cols)] for _ in range(table.rows)]
                
                for cell in table.cells:
                    if 0 <= cell.row < table.rows and 0 <= cell.col < table.cols:
                        grid[cell.row][cell.col] = cell.text
                
                # None을 빈 문자열로 변환
                for r in range(table.rows):
                    for c in range(table.cols):
                        if grid[r][c] is None:
                            grid[r][c] = ""
                
                # 헤더/데이터 분리 (첫 행을 헤더로)
                headers = [grid[0]] if grid else []
                rows = grid[1:] if len(grid) > 1 else []
                
                # 표 제목 추출 (이전 문단에서)
                table_title = ""
                if prev_para_text and len(prev_para_text) < 100:
                    table_title = prev_para_text
                
                table_struct = TableStructure(
                    table_id=table_id,
                    title=table_title,
                    headers=headers,
                    rows=rows,
                    bbox=table_bbox,
                    page=section.index,
                    metadata={
                        "original_id": table.id,
                        "z_order": table.z_order,
                        "row_count": table.rows,
                        "col_count": table.cols,
                    }
                )
                tables.append(table_struct)
                
                # 테이블 요소도 elements에 추가
                table_elem = DocumentElement(
                    element_id=table_id,
                    element_type="table",
                    text=f"[표 {table.rows}×{table.cols}] {table_title}",
                    bbox=table_bbox,
                    page=section.index,
                    metadata=table_struct.metadata,
                )
                elements.append(table_elem)
                
                # 셀들도 개별 요소로 추가
                cell_y = table_y
                for r, row in enumerate(grid):
                    cell_x = table_x
                    for c, cell_text in enumerate(row):
                        cell_id = next_id()
                        
                        # 셀 크기 추정
                        cell_width = table_size["width_mm"] / table.cols
                        cell_height = table_size["height_mm"] / table.rows
                        
                        # 실제 셀 크기 정보가 있으면 사용
                        matching_cells = [cell for cell in table.cells if cell.row == r and cell.col == c]
                        if matching_cells:
                            cell_obj = matching_cells[0]
                            cell_size = cell_obj.size.to_mm()
                            if cell_size["width_mm"] > 0:
                                cell_width = cell_size["width_mm"]
                            if cell_size["height_mm"] > 0:
                                cell_height = cell_size["height_mm"]
                        
                        cell_elem = DocumentElement(
                            element_id=cell_id,
                            element_type="table_cell",
                            text=cell_text,
                            bbox=BoundingBox(x=cell_x, y=cell_y, width=cell_width, height=cell_height),
                            page=section.index,
                            parent_id=table_id,
                            metadata={"row": r, "col": c},
                        )
                        elements.append(cell_elem)
                        table_elem.children.append(cell_id)
                        
                        cell_x += cell_width
                    cell_y += cell_height
                
                # 테이블 다음 위치로 current_y 업데이트
                current_y = table_bbox.y + table_bbox.height + 2.0  # 테이블 후 여백 2mm
            
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
        images=[],
        metadata={
            "version": f"{doc.version.application} {doc.version.app_version}",
        }
    )


# =============================================================================
# HWP 문서 추출 함수
# =============================================================================

def extract_from_hwp(doc) -> ExtractedDocument:
    """
    HWP 문서에서 구조화된 정보 추출
    
    Args:
        doc: HwpDocument 객체 (hwp_parser_cursor에서 파싱된 문서)
    
    Returns:
        ExtractedDocument: 구조화된 문서 정보
    """
    elements = []
    tables = []
    headings = []
    paragraphs = []
    pages = []
    
    element_counter = 0
    
    def next_id() -> str:
        nonlocal element_counter
        element_counter += 1
        return f"elem_{element_counter:04d}"
    
    for section in doc.sections:
        # 페이지 정보
        page_width_mm = section.page_width * HWPUNIT_TO_MM if section.page_width else 210.0
        page_height_mm = section.page_height * HWPUNIT_TO_MM if section.page_height else 297.0
        margin_top = 4.0  # 기본 여백 (HWP 파일에서 가져올 수 있으면 사용)
        margin_bottom = 4.0
        margin_left = 4.0
        margin_right = 4.0
        
        # 콘텐츠 영역 높이
        content_height = page_height_mm - margin_top - margin_bottom
        
        # 현재 Y 위치 추적
        current_y = margin_top
        line_height = 5.0  # 기본 줄 높이 (mm)
        current_page = section.index
        
        # 페이지 정보 추가 (첫 페이지)
        page_info = PageInfo(
            page_num=current_page,
            width=page_width_mm,
            height=page_height_mm,
            margin_left=margin_left,
            margin_right=margin_right,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
        )
        pages.append(page_info)
        
        for para in section.paragraphs:
            text = para.plain_text.strip()
            
            if not text:
                continue
            
            # 텍스트 높이 추정 (줄 수 기반)
            line_count = text.count('\n') + 1
            text_height = line_count * line_height
            
            # 페이지 경계 체크 - 넘어가면 새 페이지로
            if current_y + text_height > page_height_mm - margin_bottom:
                current_page += 1
                current_y = margin_top
                
                # 새 페이지 정보 추가 (중복 방지)
                if not any(p.page_num == current_page for p in pages):
                    new_page_info = PageInfo(
                        page_num=current_page,
                        width=page_width_mm,
                        height=page_height_mm,
                        margin_left=margin_left,
                        margin_right=margin_right,
                        margin_top=margin_top,
                        margin_bottom=margin_bottom,
                    )
                    pages.append(new_page_info)
            
            # 바운딩 박스
            bbox = BoundingBox(
                x=margin_left,
                y=current_y,
                width=page_width_mm - margin_left - margin_right,
                height=text_height,
            )
            
            # 제목 판별
            is_heading = False
            heading_level = 0
            
            import re
            # 한글 가나다 패턴
            if re.match(r'^[가-힣]\.\s', text):
                is_heading = True
                heading_level = 2
            # 숫자 패턴  
            elif re.match(r'^\d+\.\s', text):
                is_heading = True
                heading_level = 2
            # 【】 패턴
            elif re.match(r'^【.+】', text):
                is_heading = True
                heading_level = 2
            # 짧은 텍스트
            elif len(text) < 30 and not text.endswith(('.', '다', '요')):
                is_heading = True
                heading_level = 1
            
            # 요소 생성
            elem_type = "heading" if is_heading else "paragraph"
            elem = DocumentElement(
                element_id=next_id(),
                element_type=elem_type,
                text=text,
                bbox=bbox,
                page=current_page,  # 현재 페이지 번호 사용
                level=heading_level,
            )
            elements.append(elem)
            
            if is_heading:
                headings.append(elem)
            else:
                paragraphs.append(elem)
            
            # 테이블 처리
            for table in para.tables:
                table_id = next_id()
                table_height = table.rows * line_height
                
                table_bbox = BoundingBox(
                    x=margin_left,
                    y=current_y + text_height,
                    width=page_width_mm - margin_left - margin_right,
                    height=table_height,
                )
                
                # 셀 데이터
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
                
                table_struct = TableStructure(
                    table_id=table_id,
                    title="",
                    headers=headers,
                    rows=rows,
                    bbox=table_bbox,
                    page=current_page,  # 현재 페이지 번호 사용
                )
                tables.append(table_struct)
            
            current_y += text_height + 2.0  # 문단 간격
    
    return ExtractedDocument(
        title=doc.title,
        source_file=str(doc.file_path),
        file_type="hwp",
        pages=pages,
        elements=elements,
        tables=tables,
        headings=headings,
        paragraphs=paragraphs,
        images=[],
        metadata={
            "version": doc.header.version,
            "is_compressed": doc.header.is_compressed,
        }
    )


def extract_document_with_images(
    doc,
    extract_images: bool = True,
    save_images_dir: Optional[Path] = None,
) -> ExtractedDocument:
    """
    문서에서 구조화된 정보와 이미지를 함께 추출
    """
    extracted = extract_document_elements(doc)
    
    if extract_images and HAS_IMAGE_EXTRACTOR:
        file_path = Path(extracted.source_file)
        
        try:
            if extracted.file_type == "hwp":
                images = extract_images_from_hwp(file_path)
            elif extracted.file_type == "hwpx":
                images = extract_images_from_hwpx(file_path)
            else:
                images = []
            
            for img in images:
                saved_path = ""
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

def extract_document_elements(doc) -> ExtractedDocument:
    """
    문서에서 구조화된 정보를 추출합니다.
    
    HWPX와 HWP 문서를 자동으로 감지하여 적절한 추출 함수를 호출합니다.
    
    Args:
        doc: HwpxDocument 또는 HwpDocument 객체
    
    Returns:
        ExtractedDocument: 구조화된 문서 정보
    
    사용 예시:
        from hwpx_parser_cursor import parse_hwpx_file
        from document_extractor import extract_document_elements
        
        doc = parse_hwpx_file("document.hwpx")
        extracted = extract_document_elements(doc)
        
        # JSON으로 저장
        with open("extracted.json", "w") as f:
            f.write(extracted.to_json())
        
        # 구조화된 텍스트
        print(extracted.to_structured_text())
    """
    # 문서 타입 감지
    if hasattr(doc, 'sections') and hasattr(doc, 'version'):
        # HWPX 문서
        if hasattr(doc.sections[0], 'page_props'):
            return extract_from_hwpx(doc)
    
    if hasattr(doc, 'header') and hasattr(doc.header, 'is_compressed'):
        # HWP 문서
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
    show_type_colors: bool = True,
    scale: float = 3.0,
    font_size: int = 10,
) -> Path:
    """
    추출된 문서 요소를 시각화합니다.
    
    Args:
        extracted: ExtractedDocument 객체
        output_path: 출력 이미지 경로
        page_num: 표시할 페이지 번호
        show_bbox: 바운딩 박스 표시 여부
        show_text: 텍스트 표시 여부
        show_type_colors: 요소 유형별 색상 구분
        scale: 확대 비율 (1mm = scale 픽셀)
        font_size: 폰트 크기
    
    Returns:
        Path: 저장된 이미지 경로
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        raise ImportError("Pillow 라이브러리가 필요합니다. pip install Pillow")
    
    if page_num >= len(extracted.pages):
        raise ValueError(f"페이지 {page_num}이 존재하지 않습니다.")
    
    page = extracted.pages[page_num]
    page_elements = [e for e in extracted.elements if e.page == page_num]
    
    # 요소들의 실제 범위 계산 (자동 스케일링용)
    if page_elements:
        max_y = max(e.bbox.y + e.bbox.height for e in page_elements)
        min_y = min(e.bbox.y for e in page_elements)
    else:
        max_y = page.height
        min_y = 0
    
    # 페이지 범위를 초과하면 Y 좌표 스케일링 비율 계산
    # 범례 영역(80px/scale ≈ 27mm)을 제외한 가용 높이
    legend_space = 30  # mm
    available_height = page.height - page.margin_top - legend_space
    content_height = max_y - min_y
    
    # Y 스케일 비율 (내용이 페이지를 초과하면 축소)
    if content_height > available_height and content_height > 0:
        y_scale_factor = available_height / content_height
    else:
        y_scale_factor = 1.0
    
    # 이미지 크기
    img_width = int(page.width * scale)
    img_height = int(page.height * scale)
    
    # 이미지 생성
    img = Image.new('RGB', (img_width, img_height), 'white')
    draw = ImageDraw.Draw(img)
    
    # 폰트
    try:
        font = ImageFont.truetype("/System/Library/Fonts/AppleSDGothicNeo.ttc", font_size)
    except:
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", font_size)
        except:
            font = ImageFont.load_default()
    
    # 색상 및 약어 정의
    colors = {
        "heading": {"outline": "#E91E63", "fill": "#FCE4EC", "abbr": "h"},  # 분홍
        "paragraph": {"outline": "#2196F3", "fill": "#E3F2FD", "abbr": "p"},  # 파랑
        "table": {"outline": "#4CAF50", "fill": "#E8F5E9", "abbr": "t"},  # 녹색
        "table_cell": {"outline": "#FF9800", "fill": "#FFF3E0", "abbr": "c"},  # 주황
        "text": {"outline": "#9C27B0", "fill": "#F3E5F5", "abbr": "x"},  # 보라
    }
    
    # 페이지 테두리
    draw.rectangle([(0, 0), (img_width - 1, img_height - 1)], outline='black', width=2)
    
    # 여백 영역 (본문 영역)
    content_top = int(page.margin_top * scale)
    content_bottom = int((page.height - page.margin_bottom) * scale)
    content_left = int(page.margin_left * scale)
    content_right = int((page.width - page.margin_right) * scale)
    draw.rectangle([(content_left, content_top), (content_right, content_bottom)], outline='lightgray', width=1)
    
    # 테이블 영역 및 셀 텍스트 수집 (중복 paragraph 제거용)
    table_regions = []
    table_cell_texts = set()  # 테이블 셀 텍스트 모음
    
    for elem in page_elements:
        if elem.element_type == "table":
            table_regions.append({
                "x1": elem.bbox.x,
                "y1": elem.bbox.y,
                "x2": elem.bbox.x + elem.bbox.width,
                "y2": elem.bbox.y + elem.bbox.height,
            })
        elif elem.element_type == "table_cell":
            # 테이블 셀 텍스트 수집 (공백 제거하고 비교용)
            cell_text = elem.text.strip()
            if cell_text:
                table_cell_texts.add(cell_text)
    
    def is_inside_table(bbox):
        """주어진 bbox가 테이블 영역 안에 있는지 확인"""
        for tr in table_regions:
            if (bbox.x >= tr["x1"] - 1 and bbox.x + bbox.width <= tr["x2"] + 1 and
                bbox.y >= tr["y1"] - 1 and bbox.y + bbox.height <= tr["y2"] + 1):
                return True
        return False
    
    def is_duplicate_cell_text(text):
        """텍스트가 테이블 셀 내용과 중복인지 확인"""
        text_stripped = text.strip()
        if text_stripped in table_cell_texts:
            return True
        # 부분 매칭 (텍스트가 셀 텍스트를 포함하거나 그 반대)
        for cell_text in table_cell_texts:
            if cell_text in text_stripped or text_stripped in cell_text:
                if len(text_stripped) > 3 and len(cell_text) > 3:  # 짧은 텍스트 제외
                    return True
        return False
    
    # 범례 영역 높이 계산 (80px 확보)
    legend_height = 80
    max_content_y = img_height - legend_height
    
    # 최소 스케일 비율 설정 (너무 압축되지 않도록)
    min_scale_factor = 0.3  # 최소 30%까지만 축소
    y_scale_factor = max(y_scale_factor, min_scale_factor)
    
    # 요소들 그리기 (Y 스케일링 적용, 중복 제거)
    for elem in page_elements:
        # 테이블 내부의 paragraph 또는 셀 텍스트와 중복된 paragraph는 스킵
        if elem.element_type == "paragraph":
            if is_inside_table(elem.bbox):
                continue
            if is_duplicate_cell_text(elem.text):
                continue
        
        # 테이블 셀은 표시하지 않음 (테이블만 표시하여 깔끔하게)
        if elem.element_type == "table_cell":
            continue
        
        color = colors.get(elem.element_type, colors["text"])
        
        # Y 좌표 스케일링 적용
        scaled_y = (elem.bbox.y - min_y) * y_scale_factor + page.margin_top
        scaled_height = elem.bbox.height * y_scale_factor
        
        # 좌표 변환
        x1 = max(0, int(elem.bbox.x * scale))
        y1 = max(0, int(scaled_y * scale))
        x2 = min(img_width - 1, int((elem.bbox.x + elem.bbox.width) * scale))
        y2 = min(max_content_y - 5, int((scaled_y + scaled_height) * scale))
        
        # 너무 작거나 범위 밖이면 스킵
        if x2 <= x1 or y2 <= y1:
            continue
        
        if show_bbox:
            if show_type_colors:
                draw.rectangle([(x1, y1), (x2, y2)], outline=color["outline"], width=1)
            else:
                draw.rectangle([(x1, y1), (x2, y2)], outline='blue', width=1)
        
        if show_text and elem.text.strip():
            display_text = elem.text.strip()
            # 박스 너비에 맞게 텍스트 길이 제한
            box_width = x2 - x1
            max_chars = max(5, int(box_width / 6))  # 대략 글자당 6px
            if len(display_text) > max_chars:
                display_text = display_text[:max_chars - 3] + "..."
            
            try:
                # 요소 유형 약어 표시 (h:, p:, t:, c: 등)
                abbr = color.get("abbr", "?")
                draw.text((x1 + 2, y1 + 2), f"{abbr}:{display_text}", fill='black', font=font)
            except:
                pass
    
    # 제목
    title = f"{extracted.title} - Page {page_num + 1}/{len(extracted.pages)}"
    draw.text((10, 5), title, fill='black', font=font)
    
    # 범례 (약어와 함께 표시)
    legend_y = img_height - 80
    draw.text((10, legend_y), "범례:", fill='black', font=font)
    legend_y += 15
    for elem_type, color_info in colors.items():
        abbr = color_info.get("abbr", "?")
        draw.rectangle([(10, legend_y), (25, legend_y + 12)], fill=color_info["fill"], outline=color_info["outline"])
        draw.text((30, legend_y), f"{abbr} - {elem_type}", fill='black', font=font)
        legend_y += 15
    
    # 저장
    output_path = Path(output_path)
    img.save(output_path)
    print(f"✅ 시각화 저장: {output_path}")
    
    return output_path


def create_visualization_report(
    extracted: ExtractedDocument,
    output_dir: str | Path,
) -> list[Path]:
    """
    전체 문서에 대한 시각화 리포트 생성
    
    Args:
        extracted: ExtractedDocument 객체
        output_dir: 출력 디렉토리
    
    Returns:
        list[Path]: 생성된 파일 경로 리스트
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    saved_files = []
    
    # 각 페이지 시각화
    for page_num in range(len(extracted.pages)):
        img_path = output_dir / f"{extracted.title}_page_{page_num + 1:03d}.png"
        visualize_elements(extracted, img_path, page_num=page_num)
        saved_files.append(img_path)
    
    # JSON 저장
    json_path = output_dir / f"{extracted.title}_extracted.json"
    with open(json_path, "w", encoding="utf-8") as f:
        f.write(extracted.to_json())
    saved_files.append(json_path)
    print(f"✅ JSON 저장: {json_path}")
    
    # 구조화된 텍스트 저장
    txt_path = output_dir / f"{extracted.title}_structured.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(extracted.to_structured_text())
    saved_files.append(txt_path)
    print(f"✅ 구조화된 텍스트 저장: {txt_path}")
    
    # 표 요약 저장
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
        print(f"✅ 표 목록 저장: {tables_path}")
    
    # 클래스 약어 설명 저장
    classes_path = output_dir / "classes.md"
    with open(classes_path, "w", encoding="utf-8") as f:
        f.write("# Element Classes / 요소 클래스\n\n")
        f.write("시각화에서 사용되는 요소 유형의 약어와 설명입니다.\n\n")
        f.write("| 약어 | 클래스명 | 설명 | 색상 |\n")
        f.write("|:----:|:---------|:-----|:-----|\n")
        f.write("| `h` | heading | 제목 (가. 나. 다. 또는 1. 2. 3. 패턴) | 🟪 분홍 (#E91E63) |\n")
        f.write("| `p` | paragraph | 일반 문단 텍스트 | 🟦 파랑 (#2196F3) |\n")
        f.write("| `t` | table | 표 (테이블 전체) | 🟩 녹색 (#4CAF50) |\n")
        f.write("| `c` | table_cell | 표 셀 (개별 셀) | 🟧 주황 (#FF9800) |\n")
        f.write("| `x` | text | 기타 텍스트 | 🟪 보라 (#9C27B0) |\n")
        f.write("\n## 시각화 예시\n\n")
        f.write("```\n")
        f.write("h:가. 광고심의신청 접수정보    → 제목\n")
        f.write("p:은행명                       → 문단\n")
        f.write("t:[표 3×4] 접수정보            → 표\n")
        f.write("c:신청자                       → 표 셀\n")
        f.write("```\n\n")
        f.write("## JSON/Markdown 출력\n\n")
        f.write("- `element_type` 필드에 전체 클래스명이 저장됩니다\n")
        f.write("- 예: `\"element_type\": \"heading\"`\n")
    saved_files.append(classes_path)
    print(f"✅ 클래스 설명 저장: {classes_path}")
    
    # 이미지 목록 저장 (OCR 연동용)
    if extracted.images:
        images_json_path = output_dir / f"{extracted.title}_images.json"
        images_data = {
            "document_title": extracted.title,
            "image_count": len(extracted.images),
            "images": [img.to_ocr_dict() for img in extracted.images],
        }
        with open(images_json_path, "w", encoding="utf-8") as f:
            json.dump(images_data, f, ensure_ascii=False, indent=2)
        saved_files.append(images_json_path)
        print(f"✅ 이미지 메타데이터 저장: {images_json_path}")
        
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
        saved_files.append(images_md_path)
        print(f"✅ 이미지 목록 저장: {images_md_path}")
    
    return saved_files


# =============================================================================
# 테스트
# =============================================================================

if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).parent))
    
    from hwpx_parser_cursor import parse_hwpx
    from hwp_parser_cursor import parse_hwp
    
    # 테스트 파일
    data_dir = Path(__file__).parent.parent / "data" / "docs"
    output_dir = Path(__file__).parent / "results"
    
    hwpx_file = data_dir / "은행권 광고심의 결과 보고서(양식)vF (1).hwpx"
    hwp_file = data_dir / "2. [농협] 광고안(B).hwp"
    
    print("=" * 70)
    print("📄 Document Extractor 테스트")
    print("=" * 70)
    
    # HWPX 테스트
    if hwpx_file.exists():
        print(f"\n🔍 HWPX 파일 처리: {hwpx_file.name}")
        doc = parse_hwpx(hwpx_file)
        
        # 이미지 저장 디렉토리
        hwpx_img_dir = output_dir / "hwpx_extracted" / "images"
        extracted = extract_document_with_images(doc, extract_images=True, save_images_dir=hwpx_img_dir)
        
        print(f"   - 요소 수: {len(extracted.elements)}")
        print(f"   - 제목 수: {len(extracted.headings)}")
        print(f"   - 표 수: {len(extracted.tables)}")
        print(f"   - 이미지 수: {len(extracted.images)}")
        
        # 좌표 확인
        print(f"\n   📍 좌표 샘플 (처음 5개 요소):")
        for elem in extracted.elements[:5]:
            print(f"      {elem.element_type}: ({elem.bbox.x:.1f}, {elem.bbox.y:.1f}) {elem.bbox.width:.1f}×{elem.bbox.height:.1f}mm")
            text_preview = elem.text[:30] if len(elem.text) > 30 else elem.text
            print(f"         텍스트: {text_preview}...")
        
        # 시각화
        create_visualization_report(extracted, output_dir / "hwpx_extracted")
    
    # HWP 테스트
    if hwp_file.exists():
        print(f"\n🔍 HWP 파일 처리: {hwp_file.name}")
        doc = parse_hwp(hwp_file)
        
        # 이미지 저장 디렉토리
        hwp_img_dir = output_dir / "hwp_extracted" / "images"
        extracted = extract_document_with_images(doc, extract_images=True, save_images_dir=hwp_img_dir)
        
        print(f"   - 요소 수: {len(extracted.elements)}")
        print(f"   - 제목 수: {len(extracted.headings)}")
        print(f"   - 표 수: {len(extracted.tables)}")
        print(f"   - 이미지 수: {len(extracted.images)}")
        
        # 좌표 확인
        print(f"\n   📍 좌표 샘플 (처음 5개 요소):")
        for elem in extracted.elements[:5]:
            print(f"      {elem.element_type}: ({elem.bbox.x:.1f}, {elem.bbox.y:.1f}) {elem.bbox.width:.1f}×{elem.bbox.height:.1f}mm")
            text_preview = elem.text[:30] if len(elem.text) > 30 else elem.text
            print(f"         텍스트: {text_preview}...")
        
        # 시각화
        create_visualization_report(extracted, output_dir / "hwp_extracted")
    
    print("\n✅ 테스트 완료!")

