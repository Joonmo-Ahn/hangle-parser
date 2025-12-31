"""
HWP Parser - 한글 문서 파일(.hwp) 파싱

=============================================================================
HWP 파일이란?
=============================================================================
HWP는 한글과컴퓨터의 "아래아한글"에서 사용하는 문서 형식입니다.
HWPX(XML 기반)와 달리, HWP는 OLE(Object Linking and Embedding) 
compound document 형식을 사용합니다.

HWP 파일 내부 구조:
    HWP 파일 (OLE Compound Document)
    ├── FileHeader        # 파일 헤더 (버전, 속성, 암호화 정보 등)
    ├── DocInfo           # 문서 정보 (스타일, 폰트, 문단 설정 등)
    ├── BodyText/         # 본문 텍스트
    │   ├── Section0      # 첫 번째 섹션
    │   ├── Section1      # 두 번째 섹션
    │   └── ...
    ├── BinData/          # 바이너리 데이터 (이미지 등)
    │   ├── BIN0001.jpg
    │   └── ...
    ├── PrvText           # 미리보기 텍스트
    ├── PrvImage          # 미리보기 이미지
    └── Scripts/          # 스크립트 (매크로 등)

데이터 압축:
    - 대부분의 스트림은 zlib으로 압축되어 있습니다.
    - FileHeader의 플래그로 압축 여부를 확인합니다.

레코드 구조:
    HWP의 데이터는 "레코드" 단위로 저장됩니다.
    각 레코드는 4바이트 헤더 + 데이터로 구성됩니다.
    
    [4바이트 헤더]
    - TagID (10비트): 레코드 종류
    - Level (10비트): 레코드 깊이
    - Size (12비트): 데이터 크기 (0xFFF이면 다음 4바이트가 크기)

필요한 라이브러리:
    pip install olefile

=============================================================================
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

# HWP 레코드 태그 ID (주요 태그만)
class HwpTagId:
    """HWP 레코드 태그 ID 상수"""
    # 문서 정보 관련
    DOCUMENT_PROPERTIES = 0x00  # 문서 속성
    ID_MAPPINGS = 0x01          # ID 매핑
    BIN_DATA = 0x02             # 바이너리 데이터 정보
    FACE_NAME = 0x03            # 글꼴 이름
    BORDER_FILL = 0x04          # 테두리/배경
    CHAR_SHAPE = 0x05           # 글자 모양
    TAB_DEF = 0x06              # 탭 정의
    NUMBERING = 0x07            # 번호 매기기
    BULLET = 0x08               # 글머리표
    PARA_SHAPE = 0x09           # 문단 모양
    STYLE = 0x0A                # 스타일
    
    # 본문 관련
    PARA_HEADER = 0x42          # 문단 헤더
    PARA_TEXT = 0x43            # 문단 텍스트
    PARA_CHAR_SHAPE = 0x44      # 문단 내 글자 모양
    PARA_LINE_SEG = 0x45        # 문단 라인 세그먼트
    PARA_RANGE_TAG = 0x46       # 문단 범위 태그
    CTRL_HEADER = 0x47          # 컨트롤 헤더
    LIST_HEADER = 0x48          # 리스트 헤더
    PAGE_DEF = 0x49             # 페이지 정의
    FOOTNOTE_SHAPE = 0x4A       # 각주 모양
    PAGE_BORDER_FILL = 0x4B     # 쪽 테두리/배경
    
    # 표 관련
    TABLE = 0x4D                # 표
    TABLE_CELL = 0x4E           # 표 셀


# 파일 헤더 플래그
class HwpHeaderFlag:
    """파일 헤더 속성 플래그"""
    COMPRESSED = 0x01           # 압축 여부
    ENCRYPTED = 0x02            # 암호화 여부
    DISTRIBUTE = 0x04           # 배포용 문서
    SCRIPT = 0x08               # 스크립트 저장
    DRM = 0x10                  # DRM 보안
    HAS_XML_TEMPLATE = 0x20     # XML 템플릿 스토리지
    VCS = 0x40                  # 문서 이력 정보
    HAS_ELECTRONIC_SIGN = 0x80  # 전자 서명 정보


# =============================================================================
# 데이터 클래스 정의
# =============================================================================

@dataclass
class HwpRecord:
    """
    HWP 레코드 (데이터의 기본 단위)
    
    Attributes:
        tag_id: 레코드 종류 (HwpTagId 참조)
        level: 레코드 깊이 (중첩 수준)
        size: 데이터 크기
        data: 원시 데이터
    """
    tag_id: int
    level: int
    size: int
    data: bytes


@dataclass
class CharShape:
    """글자 모양 정보"""
    font_id: int = 0            # 글꼴 ID
    font_size: int = 1000       # 글자 크기 (1/100 pt)
    bold: bool = False          # 굵게
    italic: bool = False        # 기울임
    underline: bool = False     # 밑줄
    color: int = 0              # 글자 색상


@dataclass
class ParaShape:
    """문단 모양 정보"""
    align: int = 0              # 정렬 (0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데)
    left_margin: int = 0        # 왼쪽 여백
    right_margin: int = 0       # 오른쪽 여백
    indent: int = 0             # 들여쓰기
    line_spacing: int = 160     # 줄 간격 (%)


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


@dataclass
class Table:
    """테이블 데이터"""
    rows: int
    cols: int
    cells: list[TableCell] = field(default_factory=list)
    
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


@dataclass
class Paragraph:
    """문단 데이터"""
    text: str = ""
    char_shapes: list[CharShape] = field(default_factory=list)
    para_shape: ParaShape = field(default_factory=ParaShape)
    tables: list[Table] = field(default_factory=list)
    
    @property
    def plain_text(self) -> str:
        """순수 텍스트만 반환 (제어 문자 제거)"""
        # HWP 특수 문자 제거
        result = []
        for char in self.text:
            code = ord(char)
            # 일반 문자만 포함 (특수 제어 문자 제외)
            if code >= 32 or char in '\n\t':
                result.append(char)
        return ''.join(result)


@dataclass
class Section:
    """섹션 데이터"""
    index: int
    paragraphs: list[Paragraph] = field(default_factory=list)
    page_width: int = 0         # 용지 너비 (HWPUNIT)
    page_height: int = 0        # 용지 높이 (HWPUNIT)
    
    @property
    def full_text(self) -> str:
        """섹션의 전체 텍스트"""
        return "\n".join(p.plain_text for p in self.paragraphs if p.plain_text.strip())


@dataclass
class FontInfo:
    """글꼴 정보"""
    id: int
    name: str
    type: str = "TTF"


@dataclass
class FileHeader:
    """파일 헤더 정보"""
    signature: str = ""
    version: str = ""
    flags: int = 0
    is_compressed: bool = False
    is_encrypted: bool = False


@dataclass
class HwpDocument:
    """
    HWP 문서 전체
    
    Attributes:
        file_path: 원본 파일 경로
        header: 파일 헤더 정보
        fonts: 글꼴 목록
        sections: 섹션 목록
        preview_text: 미리보기 텍스트
    """
    file_path: Path
    header: FileHeader = field(default_factory=FileHeader)
    fonts: list[FontInfo] = field(default_factory=list)
    sections: list[Section] = field(default_factory=list)
    preview_text: str = ""
    
    @property
    def title(self) -> str:
        """문서 제목 (파일명)"""
        return self.file_path.stem
    
    def to_text(self) -> str:
        """전체 텍스트 추출"""
        return "\n\n".join(s.full_text for s in self.sections if s.full_text)
    
    def to_markdown(self) -> str:
        """마크다운으로 변환"""
        lines = [f"# {self.title}", ""]
        
        # 문서 정보
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
        """JSON으로 변환"""
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
                    "page_width": s.page_width,
                    "page_height": s.page_height,
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


# =============================================================================
# HWP 파서 클래스
# =============================================================================

class HwpParser:
    """
    HWP 파일 파서
    
    사용법:
        parser = HwpParser("document.hwp")
        doc = parser.parse()
        print(doc.to_text())
    
    필요한 라이브러리:
        pip install olefile
    """
    
    def __init__(self, file_path: str | Path):
        """
        파서 초기화
        
        Args:
            file_path: HWP 파일 경로
        
        Raises:
            FileNotFoundError: 파일이 존재하지 않을 때
            ImportError: olefile 라이브러리가 설치되지 않았을 때
        """
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
        """
        HWP 파일 전체 파싱
        
        Returns:
            HwpDocument: 파싱된 문서 객체
        """
        doc = HwpDocument(file_path=self.file_path)
        
        try:
            self.ole = olefile.OleFileIO(str(self.file_path))
            
            # 1. 파일 헤더 파싱
            doc.header = self._parse_file_header()
            self.is_compressed = doc.header.is_compressed
            
            # 2. 문서 정보 파싱 (글꼴 등)
            doc.fonts = self._parse_doc_info()
            
            # 3. 본문 파싱
            doc.sections = list(self._parse_body_text())
            
            # 4. 미리보기 텍스트
            doc.preview_text = self._get_preview_text()
            
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
        
        # 시그니처 (32바이트)
        header.signature = data[:32].decode('utf-8', errors='ignore').rstrip('\x00')
        
        # 버전 (4바이트, offset 32)
        if len(data) >= 36:
            version = struct.unpack('<I', data[32:36])[0]
            major = (version >> 24) & 0xFF
            minor = (version >> 16) & 0xFF
            build = (version >> 8) & 0xFF
            revision = version & 0xFF
            header.version = f"{major}.{minor}.{build}.{revision}"
        
        # 플래그 (4바이트, offset 36)
        if len(data) >= 40:
            header.flags = struct.unpack('<I', data[36:40])[0]
            header.is_compressed = bool(header.flags & HwpHeaderFlag.COMPRESSED)
            header.is_encrypted = bool(header.flags & HwpHeaderFlag.ENCRYPTED)
        
        return header
    
    def _parse_doc_info(self) -> list[FontInfo]:
        """문서 정보 파싱 (글꼴 정보 추출)"""
        fonts = []
        
        if not self.ole.exists("DocInfo"):
            return fonts
        
        data = self._read_stream("DocInfo")
        if not data:
            return fonts
        
        # 레코드 순회하며 글꼴 정보 추출
        font_id = 0
        for record in self._iter_records(data):
            if record.tag_id == HwpTagId.FACE_NAME:
                font_name = self._decode_text(record.data)
                fonts.append(FontInfo(id=font_id, name=font_name))
                font_id += 1
        
        return fonts
    
    def _parse_body_text(self) -> Iterator[Section]:
        """본문 텍스트 파싱"""
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
        table_row = 0
        table_col = 0
        
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
            
            # 표 시작
            elif tag == HwpTagId.TABLE:
                if len(record.data) >= 8:
                    flags = struct.unpack('<I', record.data[:4])[0]
                    rows = struct.unpack('<H', record.data[4:6])[0]
                    cols = struct.unpack('<H', record.data[6:8])[0]
                    current_table = Table(rows=rows, cols=cols)
                    table_row = 0
                    table_col = 0
            
            # 표 셀
            elif tag == HwpTagId.LIST_HEADER and current_table:
                # 셀 정보 처리
                pass
            
            # 페이지 정의
            elif tag == HwpTagId.PAGE_DEF:
                if len(record.data) >= 8:
                    section.page_width = struct.unpack('<I', record.data[:4])[0]
                    section.page_height = struct.unpack('<I', record.data[4:8])[0]
        
        # 마지막 문단 추가
        if current_para and (current_para.text.strip() or current_para.tables):
            section.paragraphs.append(current_para)
        
        return section
    
    def _read_stream(self, stream_name: str) -> bytes:
        """스트림 읽기 (압축 해제 포함)"""
        if not self.ole.exists(stream_name):
            return b''
        
        data = self.ole.openstream(stream_name).read()
        
        # 압축 해제
        if self.is_compressed and data:
            try:
                data = zlib.decompress(data, -15)  # raw deflate
            except zlib.error:
                pass  # 압축되지 않은 데이터
        
        return data
    
    def _iter_records(self, data: bytes) -> Iterator[HwpRecord]:
        """레코드 순회"""
        offset = 0
        
        while offset < len(data) - 4:
            # 4바이트 헤더 읽기
            header = struct.unpack('<I', data[offset:offset+4])[0]
            
            tag_id = header & 0x3FF           # 하위 10비트
            level = (header >> 10) & 0x3FF    # 다음 10비트
            size = (header >> 20) & 0xFFF     # 상위 12비트
            
            offset += 4
            
            # 크기가 0xFFF이면 다음 4바이트가 실제 크기
            if size == 0xFFF:
                if offset + 4 > len(data):
                    break
                size = struct.unpack('<I', data[offset:offset+4])[0]
                offset += 4
            
            # 데이터 읽기
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
            # UTF-16LE로 2바이트씩 읽기
            char_code = struct.unpack('<H', data[i:i+2])[0]
            i += 2
            
            # HWP 특수 문자 처리
            if char_code < 32:
                if char_code == 0:  # 문자열 끝
                    break
                elif char_code == 1:  # 예약
                    pass
                elif char_code == 2:  # 섹션/단 정의
                    i += 8  # 추가 데이터 스킵
                elif char_code == 3:  # 필드 시작
                    i += 8
                elif char_code == 4:  # 필드 끝
                    pass
                elif char_code == 9:  # 탭
                    result.append('\t')
                elif char_code == 10:  # 줄바꿈
                    result.append('\n')
                elif char_code == 11:  # 그리기 객체/표
                    i += 8
                elif char_code == 12:  # 예약
                    pass
                elif char_code == 13:  # 문단 끝
                    result.append('\n')
                elif char_code == 14:  # 머리말/꼬리말/각주/미주
                    i += 8
                elif char_code == 15:  # 숨은 설명
                    i += 8
                elif char_code == 16:  # 강제 줄나눔
                    result.append('\n')
                elif char_code == 17:  # 하이픈
                    result.append('-')
                elif char_code == 18:  # 예약
                    pass
                elif char_code == 19:  # 예약
                    pass
                elif char_code == 20:  # 예약
                    pass
                elif char_code == 21:  # 컨트롤 객체
                    i += 8
                elif char_code == 22:  # 예약
                    pass
                elif char_code == 23:  # 책갈피/양식
                    i += 8
                elif char_code == 24:  # 덧말
                    i += 8
                elif char_code == 25:  # 예약
                    pass
                elif char_code == 26:  # 예약
                    pass
                elif char_code == 27:  # 예약
                    pass
                elif char_code == 28:  # 예약
                    pass
                elif char_code == 29:  # 예약
                    pass
                elif char_code == 30:  # 글자 겹침
                    i += 8
                elif char_code == 31:  # 예약
                    pass
            else:
                # 일반 문자
                result.append(chr(char_code))
        
        return ''.join(result)
    
    def _decode_text(self, data: bytes) -> str:
        """일반 텍스트 디코딩 (UTF-16LE)"""
        try:
            # 속성 바이트를 건너뛰고 텍스트 추출
            # 글꼴 이름은 첫 바이트가 속성
            if data and len(data) > 1:
                text_data = data[1:]  # 첫 바이트 스킵
                # 널 문자까지만 읽기
                null_pos = text_data.find(b'\x00\x00')
                if null_pos >= 0:
                    text_data = text_data[:null_pos+1]
                return text_data.decode('utf-16le', errors='ignore').rstrip('\x00')
        except:
            pass
        return ""
    
    def _get_preview_text(self) -> str:
        """미리보기 텍스트 읽기"""
        if not self.ole.exists("PrvText"):
            return ""
        
        try:
            data = self.ole.openstream("PrvText").read()
            return data.decode('utf-16le', errors='ignore').rstrip('\x00')
        except:
            return ""
    
    def get_stream_list(self) -> list[str]:
        """파일 내 모든 스트림 목록 반환"""
        if not self.ole:
            with olefile.OleFileIO(str(self.file_path)) as ole:
                return ['/'.join(entry) for entry in ole.listdir()]
        return ['/'.join(entry) for entry in self.ole.listdir()]


# =============================================================================
# 편의 함수
# =============================================================================

def parse_hwp(file_path: str | Path) -> HwpDocument:
    """HWP 파일을 파싱하는 편의 함수"""
    parser = HwpParser(file_path)
    return parser.parse()


def extract_text_from_hwp(file_path: str | Path) -> str:
    """HWP 파일에서 텍스트만 추출"""
    doc = parse_hwp(file_path)
    return doc.to_text()


# =============================================================================
# 메인 실행부
# =============================================================================

if __name__ == "__main__":
    import sys
    
    if not HAS_OLEFILE:
        print("❌ olefile 라이브러리가 필요합니다.")
        print("   설치: pip install olefile")
        sys.exit(1)
    
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = "data/AI test용/2. [농협] 광고안(B).hwp"
    
    print(f"파싱 중: {file_path}")
    print("=" * 60)
    
    try:
        # 파일 구조 확인
        parser = HwpParser(file_path)
        print("\n📁 파일 내부 구조:")
        for stream in parser.get_stream_list():
            print(f"  - {stream}")
        
        # 파싱
        doc = parser.parse()
        
        print(f"\n📄 문서: {doc.title}")
        print(f"📋 버전: {doc.header.version}")
        print(f"🔐 암호화: {'예' if doc.header.is_encrypted else '아니오'}")
        print(f"📦 압축: {'예' if doc.header.is_compressed else '아니오'}")
        print(f"📝 글꼴 수: {len(doc.fonts)}")
        print(f"📑 섹션 수: {len(doc.sections)}")
        
        for i, font in enumerate(doc.fonts[:5]):
            print(f"    - {font.name}")
        if len(doc.fonts) > 5:
            print(f"    ... 외 {len(doc.fonts) - 5}개")
        
        for section in doc.sections:
            print(f"\n--- Section {section.index + 1} ---")
            print(f"  문단 수: {len(section.paragraphs)}")
            if section.page_width and section.page_height:
                print(f"  페이지: {section.page_width} × {section.page_height} HWPUNIT")
        
        # 텍스트 출력
        print("\n" + "=" * 60)
        print("📝 추출된 텍스트 (처음 2000자):")
        print("=" * 60)
        text = doc.to_text()
        print(text[:2000] if len(text) > 2000 else text)
        
        # 파일 저장
        output_dir = Path(file_path).parent
        
        # 텍스트 저장
        txt_file = output_dir / f"{doc.title}_extracted.txt"
        with open(txt_file, "w", encoding="utf-8") as f:
            f.write(doc.to_text())
        print(f"\n✅ 텍스트 저장: {txt_file}")
        
        # JSON 저장
        json_file = output_dir / f"{doc.title}_parsed.json"
        with open(json_file, "w", encoding="utf-8") as f:
            f.write(doc.to_json())
        print(f"✅ JSON 저장: {json_file}")
        
        # 마크다운 저장
        md_file = output_dir / f"{doc.title}_parsed.md"
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(doc.to_markdown())
        print(f"✅ 마크다운 저장: {md_file}")
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()



