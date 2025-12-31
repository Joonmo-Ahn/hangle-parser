"""
HWP Parser - 한글 문서 파일(.hwp) 파싱 및 레이아웃 정보 추출

HWP 파일은 OLE Compound Document 형식입니다.
이 파서는 텍스트, 표, 레이아웃 정보를 추출합니다.

필요한 라이브러리:
    pip install olefile

사용 예시:
    from hwp_parser import parse_hwp, extract_layout_elements

    doc = parse_hwp("document.hwp")
    print(doc.to_text())

    elements = extract_layout_elements(doc)
    for elem in elements:
        print(f"{elem.text[:20]}... at ({elem.x}, {elem.y})")
"""

from __future__ import annotations
import struct
import zlib
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Iterator, BinaryIO, Any
import json

# olefile은 선택적 의존성
try:
    import olefile
    HAS_OLEFILE = True
except ImportError:
    HAS_OLEFILE = False


# =============================================================================
# 상수 정의
# =============================================================================

# HWPUNIT to mm 변환
HWPUNIT_TO_MM = 25.4 / 7200


# HWP 레코드 태그 ID
class HwpTagId:
    """HWP 레코드 태그 ID"""
    DOCUMENT_PROPERTIES = 0x00
    ID_MAPPINGS = 0x01
    BIN_DATA = 0x02
    FACE_NAME = 0x03
    BORDER_FILL = 0x04
    CHAR_SHAPE = 0x05
    TAB_DEF = 0x06
    NUMBERING = 0x07
    BULLET = 0x08
    PARA_SHAPE = 0x09
    STYLE = 0x0A

    PARA_HEADER = 0x42
    PARA_TEXT = 0x43
    PARA_CHAR_SHAPE = 0x44
    PARA_LINE_SEG = 0x45
    PARA_RANGE_TAG = 0x46
    CTRL_HEADER = 0x47
    LIST_HEADER = 0x48
    PAGE_DEF = 0x49
    FOOTNOTE_SHAPE = 0x4A
    PAGE_BORDER_FILL = 0x4B

    TABLE = 0x4D
    TABLE_CELL = 0x4E


class HwpHeaderFlag:
    """파일 헤더 플래그"""
    COMPRESSED = 0x01
    ENCRYPTED = 0x02
    DISTRIBUTE = 0x04
    SCRIPT = 0x08
    DRM = 0x10


# =============================================================================
# 데이터 클래스
# =============================================================================

@dataclass
class BoundingBox:
    """바운딩 박스 (mm 단위)"""
    x: float = 0.0
    y: float = 0.0
    width: float = 0.0
    height: float = 0.0
    page: int = 0  # 페이지 번호 (0부터 시작)

    def to_dict(self) -> dict:
        return {
            "x": round(self.x, 2),
            "y": round(self.y, 2),
            "width": round(self.width, 2),
            "height": round(self.height, 2),
            "page": self.page,
        }

    def is_valid(self) -> bool:
        return not (self.x == 0 and self.y == 0 and self.width == 0 and self.height == 0)
    
    def clip_to_page(self, page_width: float, page_height: float, 
                     margin_left: float = 0, margin_top: float = 0,
                     margin_right: float = 0, margin_bottom: float = 0) -> "BoundingBox":
        """페이지 경계 내로 바운딩 박스 클리핑"""
        content_width = page_width - margin_left - margin_right
        content_height = page_height - margin_top - margin_bottom
        
        # X 좌표 클리핑
        x = max(self.x, margin_left)
        width = min(self.width, content_width - (x - margin_left))
        
        # Y 좌표 클리핑
        y = max(self.y, margin_top)
        height = min(self.height, content_height - (y - margin_top))
        
        return BoundingBox(
            x=x, y=y, 
            width=max(width, 0), 
            height=max(height, 0),
            page=self.page
        )


@dataclass
class HwpRecord:
    """HWP 레코드"""
    tag_id: int
    level: int
    size: int
    data: bytes


@dataclass
class CharShape:
    """글자 모양"""
    font_id: int = 0
    font_size: int = 1000  # 1/100 pt
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: int = 0


@dataclass
class ParaShape:
    """문단 모양"""
    align: int = 0  # 0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데
    left_margin: int = 0
    right_margin: int = 0
    indent: int = 0
    line_spacing: int = 160


@dataclass
class LineSegment:
    """라인 세그먼트 (HWP 본문용)"""
    text_pos: int = 0       # 텍스트 시작 위치
    vert_pos: int = 0       # 수직 위치 (HWPUNIT)
    vert_size: int = 0      # 줄 높이
    text_height: int = 0
    baseline: int = 0
    spacing: int = 0
    horz_pos: int = 0       # 수평 위치
    horz_size: int = 0      # 줄 너비
    tag: int = 0            # 태그 정보

    def to_mm(self) -> dict:
        return {
            "x_mm": round(self.horz_pos * HWPUNIT_TO_MM, 2),
            "y_mm": round(self.vert_pos * HWPUNIT_TO_MM, 2),
            "width_mm": round(self.horz_size * HWPUNIT_TO_MM, 2),
            "height_mm": round(self.vert_size * HWPUNIT_TO_MM, 2),
        }


@dataclass
class TableCell:
    """테이블 셀"""
    row: int
    col: int
    text: str
    row_span: int = 1
    col_span: int = 1
    width: int = 0
    height: int = 0
    bbox: BoundingBox = field(default_factory=BoundingBox)


@dataclass
class Table:
    """테이블"""
    rows: int
    cols: int
    cells: list[TableCell] = field(default_factory=list)
    bbox: BoundingBox = field(default_factory=BoundingBox)

    def to_markdown(self) -> str:
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


@dataclass
class Paragraph:
    """문단"""
    text: str = ""
    char_shapes: list[CharShape] = field(default_factory=list)
    para_shape: ParaShape = field(default_factory=ParaShape)
    line_segments: list[LineSegment] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    bbox: BoundingBox = field(default_factory=BoundingBox)

    @property
    def plain_text(self) -> str:
        """순수 텍스트만 반환"""
        result = []
        for char in self.text:
            code = ord(char)
            if code >= 32 or char in '\n\t':
                result.append(char)
        return ''.join(result)

    def calculate_bbox(self, margin_left: float = 0, margin_top: float = 0, 
                       content_height: float = 257.0) -> BoundingBox:
        """
        바운딩 박스 계산 (개선 버전)
        
        Args:
            margin_left: 왼쪽 여백 (mm)
            margin_top: 상단 여백 (mm)
            content_height: 콘텐츠 영역 높이 (mm, 페이지 분할용)
        """
        if not self.line_segments:
            return BoundingBox()

        valid_segments = [ls for ls in self.line_segments if ls.horz_size > 0 or ls.vert_size > 0]
        if not valid_segments:
            return BoundingBox()

        min_x = min(ls.horz_pos for ls in valid_segments)
        max_x = max(ls.horz_pos + ls.horz_size for ls in valid_segments)
        min_y = min(ls.vert_pos for ls in valid_segments)
        max_y = max(ls.vert_pos + ls.vert_size for ls in valid_segments)
        
        # HWPUNIT을 mm로 변환
        min_y_mm = min_y * HWPUNIT_TO_MM
        max_y_mm = max_y * HWPUNIT_TO_MM
        
        # 페이지 번호 계산
        page_num = int(min_y_mm // content_height) if content_height > 0 else 0
        
        # 페이지 내 상대 Y 좌표
        page_relative_y = min_y_mm - (page_num * content_height)
        if page_relative_y < 0:
            page_relative_y = 0

        return BoundingBox(
            x=margin_left + min_x * HWPUNIT_TO_MM,
            y=margin_top + page_relative_y,
            width=(max_x - min_x) * HWPUNIT_TO_MM,
            height=(max_y_mm - min_y_mm),
            page=page_num,
        )


@dataclass
class Section:
    """섹션"""
    index: int
    paragraphs: list[Paragraph] = field(default_factory=list)
    page_width: int = 0
    page_height: int = 0
    margin_left: int = 0
    margin_right: int = 0
    margin_top: int = 0
    margin_bottom: int = 0

    @property
    def full_text(self) -> str:
        return "\n".join(p.plain_text for p in self.paragraphs if p.plain_text.strip())

    def page_width_mm(self) -> float:
        return self.page_width * HWPUNIT_TO_MM if self.page_width else 210.0

    def page_height_mm(self) -> float:
        return self.page_height * HWPUNIT_TO_MM if self.page_height else 297.0

    def margin_left_mm(self) -> float:
        return self.margin_left * HWPUNIT_TO_MM if self.margin_left else 20.0

    def margin_top_mm(self) -> float:
        return self.margin_top * HWPUNIT_TO_MM if self.margin_top else 20.0


@dataclass
class FontInfo:
    """글꼴 정보"""
    id: int
    name: str
    type: str = "TTF"


@dataclass
class FileHeader:
    """파일 헤더"""
    signature: str = ""
    version: str = ""
    flags: int = 0
    is_compressed: bool = False
    is_encrypted: bool = False


@dataclass
class HwpDocument:
    """HWP 문서"""
    file_path: Path
    header: FileHeader = field(default_factory=FileHeader)
    fonts: list[FontInfo] = field(default_factory=list)
    sections: list[Section] = field(default_factory=list)
    preview_text: str = ""

    @property
    def title(self) -> str:
        return self.file_path.stem

    def to_text(self) -> str:
        return "\n\n".join(s.full_text for s in self.sections if s.full_text)

    def to_markdown(self) -> str:
        lines = [f"# {self.title}", ""]
        lines.append(f"- 버전: {self.header.version}")
        lines.append(f"- 압축: {'예' if self.header.is_compressed else '아니오'}")
        lines.append("")

        for section in self.sections:
            lines.append(f"## Section {section.index + 1}")
            lines.append("")

            for para in section.paragraphs:
                text = para.plain_text.strip()
                if text:
                    lines.append(text)
                    lines.append("")

                for table in para.tables:
                    lines.append(table.to_markdown())
                    lines.append("")

        return "\n".join(lines)

    def to_json(self) -> str:
        data = {
            "title": self.title,
            "header": {
                "version": self.header.version,
                "is_compressed": self.header.is_compressed,
                "is_encrypted": self.header.is_encrypted,
            },
            "fonts": [{"id": f.id, "name": f.name} for f in self.fonts],
            "sections": [
                {
                    "index": s.index,
                    "page_width_mm": s.page_width_mm(),
                    "page_height_mm": s.page_height_mm(),
                    "paragraphs": [
                        {
                            "text": p.plain_text,
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
                        if p.plain_text.strip() or p.tables
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
            "header": asdict(self.header),
            "unit_info": {
                "description": "좌표 단위는 mm",
            },
            "sections": []
        }

        for section in self.sections:
            section_data = {
                "index": section.index,
                "page": {
                    "width_mm": section.page_width_mm(),
                    "height_mm": section.page_height_mm(),
                    "margins_mm": {
                        "left": section.margin_left_mm(),
                        "right": section.margin_right * HWPUNIT_TO_MM if section.margin_right else 20.0,
                        "top": section.margin_top_mm(),
                        "bottom": section.margin_bottom * HWPUNIT_TO_MM if section.margin_bottom else 20.0,
                    }
                },
                "paragraphs": []
            }

            for para in section.paragraphs:
                if not para.plain_text.strip() and not para.tables:
                    continue

                para_data = {
                    "text": para.plain_text,
                    "bbox": para.bbox.to_dict() if para.bbox.is_valid() else None,
                    "line_segments": [
                        ls.to_mm() for ls in para.line_segments
                    ],
                    "tables": [
                        {
                            "rows": t.rows,
                            "cols": t.cols,
                            "bbox": t.bbox.to_dict() if t.bbox.is_valid() else None,
                            "cells": [
                                {
                                    "row": c.row,
                                    "col": c.col,
                                    "text": c.text,
                                    "bbox": c.bbox.to_dict() if c.bbox.is_valid() else None,
                                }
                                for c in t.cells
                            ]
                        }
                        for t in para.tables
                    ]
                }
                section_data["paragraphs"].append(para_data)

            data["sections"].append(section_data)

        return json.dumps(data, ensure_ascii=False, indent=2)


# =============================================================================
# HWP 파서 클래스
# =============================================================================

class HwpParser:
    """HWP 파일 파서"""

    def __init__(self, file_path: str | Path):
        if not HAS_OLEFILE:
            raise ImportError(
                "olefile 라이브러리가 필요합니다.\n"
                "설치: pip install olefile"
            )

        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

        self.ole = None
        self.is_compressed = False

    def parse(self) -> HwpDocument:
        """문서 파싱"""
        doc = HwpDocument(file_path=self.file_path)

        try:
            self.ole = olefile.OleFileIO(str(self.file_path))

            doc.header = self._parse_file_header()
            self.is_compressed = doc.header.is_compressed

            doc.fonts = self._parse_doc_info()
            doc.sections = list(self._parse_body_text())
            doc.preview_text = self._get_preview_text()

            # 바운딩 박스 계산
            self._calculate_all_bboxes(doc)

        finally:
            if self.ole:
                self.ole.close()

        return doc

    def _parse_file_header(self) -> FileHeader:
        """파일 헤더 파싱"""
        header = FileHeader()

        if not self.ole.exists("FileHeader"):
            return header

        data = self.ole.openstream("FileHeader").read()

        header.signature = data[:32].decode('utf-8', errors='ignore').rstrip('\x00')

        if len(data) >= 36:
            version = struct.unpack('<I', data[32:36])[0]
            major = (version >> 24) & 0xFF
            minor = (version >> 16) & 0xFF
            build = (version >> 8) & 0xFF
            revision = version & 0xFF
            header.version = f"{major}.{minor}.{build}.{revision}"

        if len(data) >= 40:
            header.flags = struct.unpack('<I', data[36:40])[0]
            header.is_compressed = bool(header.flags & HwpHeaderFlag.COMPRESSED)
            header.is_encrypted = bool(header.flags & HwpHeaderFlag.ENCRYPTED)

        return header

    def _parse_doc_info(self) -> list[FontInfo]:
        """문서 정보 (글꼴) 파싱"""
        fonts = []

        if not self.ole.exists("DocInfo"):
            return fonts

        data = self._read_stream("DocInfo")
        if not data:
            return fonts

        font_id = 0
        for record in self._iter_records(data):
            if record.tag_id == HwpTagId.FACE_NAME:
                font_name = self._decode_text(record.data)
                fonts.append(FontInfo(id=font_id, name=font_name))
                font_id += 1

        return fonts

    def _parse_body_text(self) -> Iterator[Section]:
        """본문 파싱"""
        section_idx = 0

        while True:
            stream_name = f"BodyText/Section{section_idx}"
            if not self.ole.exists(stream_name):
                break

            section = self._parse_section(stream_name, section_idx)
            yield section
            section_idx += 1

    def _parse_section(self, stream_name: str, index: int) -> Section:
        """단일 섹션 파싱"""
        section = Section(index=index)

        data = self._read_stream(stream_name)
        if not data:
            return section

        current_para = None
        current_table = None
        in_table = False

        for record in self._iter_records(data):
            tag = record.tag_id

            # 문단 헤더
            if tag == HwpTagId.PARA_HEADER:
                if current_para and (current_para.text.strip() or current_para.tables):
                    section.paragraphs.append(current_para)
                current_para = Paragraph()

            # 문단 텍스트
            elif tag == HwpTagId.PARA_TEXT and current_para:
                text = self._decode_para_text(record.data)
                current_para.text += text

            # 라인 세그먼트
            elif tag == HwpTagId.PARA_LINE_SEG and current_para:
                segments = self._parse_line_segments(record.data)
                current_para.line_segments.extend(segments)

            # 표
            elif tag == HwpTagId.TABLE:
                if len(record.data) >= 8:
                    rows = struct.unpack('<H', record.data[4:6])[0]
                    cols = struct.unpack('<H', record.data[6:8])[0]
                    current_table = Table(rows=rows, cols=cols)
                    in_table = True

            # 리스트 헤더 (셀 시작)
            elif tag == HwpTagId.LIST_HEADER and current_table:
                pass  # 셀 시작 처리

            # 페이지 정의
            elif tag == HwpTagId.PAGE_DEF:
                if len(record.data) >= 40:
                    section.page_width = struct.unpack('<I', record.data[:4])[0]
                    section.page_height = struct.unpack('<I', record.data[4:8])[0]
                    section.margin_left = struct.unpack('<I', record.data[8:12])[0]
                    section.margin_right = struct.unpack('<I', record.data[12:16])[0]
                    section.margin_top = struct.unpack('<I', record.data[16:20])[0]
                    section.margin_bottom = struct.unpack('<I', record.data[20:24])[0]

        # 마지막 문단
        if current_para and (current_para.text.strip() or current_para.tables):
            section.paragraphs.append(current_para)

        return section

    def _parse_line_segments(self, data: bytes) -> list[LineSegment]:
        """라인 세그먼트 파싱"""
        segments = []

        # 각 라인 세그먼트는 32바이트
        segment_size = 32
        count = len(data) // segment_size

        for i in range(count):
            offset = i * segment_size
            seg_data = data[offset:offset + segment_size]

            if len(seg_data) < segment_size:
                break

            seg = LineSegment(
                text_pos=struct.unpack('<I', seg_data[0:4])[0],
                vert_pos=struct.unpack('<i', seg_data[4:8])[0],
                vert_size=struct.unpack('<i', seg_data[8:12])[0],
                text_height=struct.unpack('<i', seg_data[12:16])[0],
                baseline=struct.unpack('<i', seg_data[16:20])[0],
                spacing=struct.unpack('<i', seg_data[20:24])[0],
                horz_pos=struct.unpack('<i', seg_data[24:28])[0],
                horz_size=struct.unpack('<i', seg_data[28:32])[0],
            )
            segments.append(seg)

        return segments

    def _read_stream(self, stream_name: str) -> bytes:
        """스트림 읽기 (압축 해제 포함)"""
        if not self.ole.exists(stream_name):
            return b''

        data = self.ole.openstream(stream_name).read()

        if self.is_compressed and data:
            try:
                data = zlib.decompress(data, -15)
            except zlib.error:
                pass

        return data

    def _iter_records(self, data: bytes) -> Iterator[HwpRecord]:
        """레코드 순회"""
        offset = 0

        while offset < len(data) - 4:
            header = struct.unpack('<I', data[offset:offset+4])[0]

            tag_id = header & 0x3FF
            level = (header >> 10) & 0x3FF
            size = (header >> 20) & 0xFFF

            offset += 4

            if size == 0xFFF:
                if offset + 4 > len(data):
                    break
                size = struct.unpack('<I', data[offset:offset+4])[0]
                offset += 4

            if offset + size > len(data):
                break

            record_data = data[offset:offset+size]
            offset += size

            yield HwpRecord(tag_id=tag_id, level=level, size=size, data=record_data)

    def _decode_para_text(self, data: bytes) -> str:
        """문단 텍스트 디코딩"""
        if not data:
            return ""

        result = []
        i = 0

        while i < len(data) - 1:
            char_code = struct.unpack('<H', data[i:i+2])[0]
            i += 2

            if char_code < 32:
                if char_code == 0:
                    break
                elif char_code == 9:
                    result.append('\t')
                elif char_code in (10, 13, 16):
                    result.append('\n')
                elif char_code == 17:
                    result.append('-')
                elif char_code in (2, 3, 11, 14, 15, 21, 23, 24, 30):
                    i += 8  # 추가 데이터 스킵
            else:
                result.append(chr(char_code))

        return ''.join(result)

    def _decode_text(self, data: bytes) -> str:
        """일반 텍스트 디코딩"""
        try:
            if data and len(data) > 1:
                text_data = data[1:]
                null_pos = text_data.find(b'\x00\x00')
                if null_pos >= 0:
                    text_data = text_data[:null_pos+1]
                return text_data.decode('utf-16le', errors='ignore').rstrip('\x00')
        except:
            pass
        return ""

    def _get_preview_text(self) -> str:
        """미리보기 텍스트"""
        if not self.ole.exists("PrvText"):
            return ""

        try:
            data = self.ole.openstream("PrvText").read()
            return data.decode('utf-16le', errors='ignore').rstrip('\x00')
        except:
            return ""

    def _calculate_all_bboxes(self, doc: HwpDocument):
        """
        모든 바운딩 박스 계산 (개선 버전)

        HWP 파일의 좌표는 라인 세그먼트의 vert_pos가 섹션 시작점 기준 절대 좌표입니다.
        페이지 높이를 기준으로 페이지를 분리하고, 각 페이지 내에서의 상대 좌표를 계산합니다.
        
        개선 사항:
        1. 페이지 경계를 고려한 좌표 계산
        2. 라인 세그먼트의 vert_pos를 절대 좌표로 직접 사용
        3. 페이지 경계 초과 시 자동 페이지 분할
        """
        for section in doc.sections:
            margin_left = section.margin_left_mm()
            margin_top = section.margin_top_mm()
            margin_bottom = section.margin_bottom * HWPUNIT_TO_MM if section.margin_bottom else 20.0
            page_height = section.page_height_mm()
            
            # 콘텐츠 영역 높이 (여백 제외)
            content_height = page_height - margin_top - margin_bottom
            
            for para in section.paragraphs:
                if not para.line_segments:
                    para.bbox = BoundingBox()
                    continue

                # 유효한 라인 세그먼트 필터링
                valid_segs = [ls for ls in para.line_segments 
                             if ls.horz_size > 0 or ls.vert_size > 0]
                if not valid_segs:
                    para.bbox = BoundingBox()
                    continue

                # 라인 세그먼트에서 좌표 추출 (HWPUNIT -> mm)
                min_x = min(ls.horz_pos for ls in valid_segs) * HWPUNIT_TO_MM
                max_x = max(ls.horz_pos + ls.horz_size for ls in valid_segs) * HWPUNIT_TO_MM
                min_y = min(ls.vert_pos for ls in valid_segs) * HWPUNIT_TO_MM
                max_y = max((ls.vert_pos + ls.vert_size) for ls in valid_segs) * HWPUNIT_TO_MM

                # 페이지 번호 계산 (0부터 시작)
                page_num = int(min_y // content_height) if content_height > 0 else 0
                
                # 페이지 내 상대 Y 좌표 계산
                page_relative_y = min_y - (page_num * content_height)
                
                # 페이지 경계 검증 및 조정
                if page_relative_y < 0:
                    page_relative_y = 0
                if page_relative_y > content_height:
                    page_relative_y = page_relative_y % content_height if content_height > 0 else 0

                # 바운딩 박스 설정 (마진 적용)
                para.bbox = BoundingBox(
                    x=margin_left + min_x,
                    y=margin_top + page_relative_y,
                    width=max(max_x - min_x, 1.0),  # 최소 너비 1mm
                    height=max(max_y - min_y, 1.0),  # 최소 높이 1mm
                )
                
                # 페이지 경계 초과 시 높이 조정
                max_allowed_height = content_height - page_relative_y
                if para.bbox.height > max_allowed_height and max_allowed_height > 0:
                    para.bbox.height = max_allowed_height
                
                # 페이지 정보 저장 (메타데이터로 활용 가능)
                para.bbox.page = page_num  # type: ignore

    def get_stream_list(self) -> list[str]:
        """파일 내 스트림 목록"""
        if not self.ole:
            with olefile.OleFileIO(str(self.file_path)) as ole:
                return ['/'.join(entry) for entry in ole.listdir()]
        return ['/'.join(entry) for entry in self.ole.listdir()]


# =============================================================================
# 레이아웃 추출 함수
# =============================================================================

@dataclass
class LayoutElement:
    """레이아웃 요소"""
    element_type: str
    text: str
    x: float
    y: float
    width: float
    height: float
    page: int = 0
    section: int = 0
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


def extract_layout_elements(doc: HwpDocument) -> tuple[list[LayoutElement], list[PageInfo]]:
    """
    문서에서 레이아웃 요소 추출

    Args:
        doc: 파싱된 HWP 문서

    Returns:
        tuple: (레이아웃 요소 리스트, 페이지 정보 리스트)
    """
    elements = []
    pages = []

    for section in doc.sections:
        page_info = PageInfo(
            page_num=section.index,
            width=section.page_width_mm(),
            height=section.page_height_mm(),
            margin_top=section.margin_top_mm(),
            margin_bottom=section.margin_bottom * HWPUNIT_TO_MM if section.margin_bottom else 20.0,
            margin_left=section.margin_left_mm(),
            margin_right=section.margin_right * HWPUNIT_TO_MM if section.margin_right else 20.0,
        )
        pages.append(page_info)

        for para in section.paragraphs:
            text = para.plain_text.strip()
            if not text and not para.tables:
                continue

            # 문단 요소
            if text and para.bbox.is_valid():
                elem = LayoutElement(
                    element_type="text",
                    text=text,
                    x=para.bbox.x,
                    y=para.bbox.y,
                    width=para.bbox.width,
                    height=para.bbox.height,
                    page=section.index,
                    section=section.index,
                    metadata={"line_count": len(para.line_segments)}
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
                        metadata={"rows": table.rows, "cols": table.cols}
                    )
                    elements.append(table_elem)

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
                                metadata={"row": cell.row, "col": cell.col}
                            )
                            elements.append(cell_elem)

    return elements, pages


def extract_layout_summary(doc: HwpDocument) -> dict:
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

def parse_hwp(file_path: str | Path) -> HwpDocument:
    """HWP 파일 파싱"""
    parser = HwpParser(file_path)
    return parser.parse()


def extract_text_from_hwp(file_path: str | Path) -> str:
    """HWP 파일에서 텍스트만 추출"""
    doc = parse_hwp(file_path)
    return doc.to_text()


# =============================================================================
# 메인
# =============================================================================

if __name__ == "__main__":
    import sys

    if not HAS_OLEFILE:
        print("olefile 라이브러리가 필요합니다.")
        print("설치: pip install olefile")
        sys.exit(1)

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        print("사용법: python hwp_parser.py <hwp_file>")
        sys.exit(1)

    print(f"파싱 중: {file_path}")
    print("=" * 60)

    try:
        parser = HwpParser(file_path)
        doc = parser.parse()

        print(f"\n📄 문서: {doc.title}")
        print(f"📋 버전: {doc.header.version}")
        print(f"📦 압축: {'예' if doc.header.is_compressed else '아니오'}")
        print(f"📑 섹션 수: {len(doc.sections)}")
        print(f"📝 글꼴 수: {len(doc.fonts)}")

        for section in doc.sections:
            print(f"\n--- Section {section.index + 1} ---")
            print(f"  페이지: {section.page_width_mm():.1f}mm × {section.page_height_mm():.1f}mm")
            print(f"  문단 수: {len(section.paragraphs)}")

        # 레이아웃 요소
        elements, pages = extract_layout_elements(doc)
        print(f"\n📐 레이아웃 요소: {len(elements)}개")

        for elem in elements[:5]:
            print(f"  - {elem.element_type}: ({elem.x:.1f}, {elem.y:.1f}) {elem.width:.1f}×{elem.height:.1f}mm")
            if elem.text:
                print(f"    텍스트: {elem.text[:50]}...")

    except Exception as e:
        print(f"\n오류: {e}")
        import traceback
        traceback.print_exc()
