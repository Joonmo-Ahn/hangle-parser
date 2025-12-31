"""
HWP/HWPX 이미지 추출 모듈

이 모듈은 HWP 및 HWPX 파일에서 임베디드 이미지를 추출합니다.

주요 기능:
    - HWP/HWPX 파일에서 이미지 추출
    - 이미지 위치 좌표 (x, y, width, height) 추출
    - WMF/EMF 벡터 형식을 PNG로 변환
    - 외부 OCR 연동을 위한 JSON 메타데이터 출력

사용 예시:
    from image_extractor import extract_images_from_hwp, extract_images_from_hwpx
    
    # HWP 파일에서 이미지 추출
    images = extract_images_from_hwp("document.hwp")
    for img in images:
        print(f"{img.filename}: {len(img.data)} bytes")
        img.save("output/")
    
    # 외부 OCR 연동용 JSON 저장
    save_images_for_ocr(images, "output/", "document_images.json")
"""

from __future__ import annotations
import struct
import zlib
import zipfile
import tempfile
import shutil
import json
import subprocess
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional, Any
import xml.etree.ElementTree as ET

# olefile은 선택적 의존성
try:
    import olefile
    HAS_OLEFILE = True
except ImportError:
    HAS_OLEFILE = False

# PIL은 선택적 의존성 (이미지 크기 확인용)
try:
    from PIL import Image
    import io
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


# HWPUNIT to mm 변환
HWPUNIT_TO_MM = 25.4 / 7200


@dataclass
class EmbeddedImage:
    """임베디드 이미지 데이터 클래스"""
    bin_id: str = ""              # BIN0001
    filename: str = ""            # BIN0001.jpg
    format: str = ""              # jpg, png, bmp
    data: bytes = b""             # 원본 바이너리 데이터
    
    # 이미지 크기 (pixels)
    pixel_width: int = 0
    pixel_height: int = 0
    
    # 문서 내 위치 (mm)
    x: float = 0.0
    y: float = 0.0
    width: float = 0.0
    height: float = 0.0
    page: int = 0
    
    # 추가 메타데이터
    compressed: bool = False
    original_size: int = 0
    element_type: str = "image"   # 요소 타입 (OCR 연동용)
    gso_type: str = ""            # GSO 컨트롤 타입 (picture, rectangle, etc.)
    z_order: int = 0              # Z 순서 (레이어 순서)
    
    def save(self, output_dir: str | Path, convert_vector: bool = True) -> Path:
        """
        이미지를 파일로 저장
        
        Args:
            output_dir: 저장 디렉토리
            convert_vector: WMF/EMF를 PNG로 변환할지 여부
            
        Returns:
            저장된 파일 경로
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # WMF/EMF 변환
        if convert_vector and self.format in ['wmf', 'emf']:
            converted_data = convert_vector_to_png(self.data, self.format)
            if converted_data:
                self.data = converted_data
                self.format = 'png'
                self.filename = f"{self.bin_id}.png"
                # 변환 후 크기 업데이트
                if HAS_PIL:
                    self.pixel_width, self.pixel_height = self.get_size_from_data()
        
        output_path = output_dir / self.filename
        with open(output_path, 'wb') as f:
            f.write(self.data)
        
        return output_path
    
    def get_size_from_data(self) -> tuple[int, int]:
        """이미지 데이터에서 크기 추출"""
        if not HAS_PIL or not self.data:
            return (0, 0)
        
        try:
            img = Image.open(io.BytesIO(self.data))
            return img.size
        except Exception:
            return (0, 0)
    
    def to_dict(self) -> dict:
        """딕셔너리로 변환 (JSON 직렬화용)"""
        return {
            "bin_id": self.bin_id,
            "filename": self.filename,
            "format": self.format,
            "element_type": self.element_type,
            "size_bytes": len(self.data),
            "pixel_width": self.pixel_width,
            "pixel_height": self.pixel_height,
            "bbox": {
                "x": round(self.x, 2),
                "y": round(self.y, 2),
                "width": round(self.width, 2),
                "height": round(self.height, 2),
                "x2": round(self.x + self.width, 2),
                "y2": round(self.y + self.height, 2),
            },
            "page": self.page,
            "compressed": self.compressed,
            "gso_type": self.gso_type,
            "z_order": self.z_order,
        }
    
    def to_ocr_dict(self) -> dict:
        """
        외부 OCR 연동용 딕셔너리
        
        이 형식은 외부 OCR 서비스에서 이미지 영역을 인식하고
        결과를 매핑하는데 사용됩니다.
        """
        return {
            "image_id": self.bin_id,
            "filename": self.filename,
            "format": self.format,
            "class": self.element_type,  # "image", "chart", "diagram", etc.
            "bbox_mm": {
                "x": round(self.x, 2),
                "y": round(self.y, 2),
                "width": round(self.width, 2),
                "height": round(self.height, 2),
            },
            "bbox_px": {
                "width": self.pixel_width,
                "height": self.pixel_height,
            },
            "page": self.page,
            "ocr_text": "",  # 외부 OCR 결과를 여기에 채움
            "ocr_confidence": 0.0,  # OCR 신뢰도
        }


# =============================================================================
# WMF/EMF 변환
# =============================================================================

def convert_vector_to_png(data: bytes, format: str, dpi: int = 300) -> Optional[bytes]:
    """
    WMF/EMF 벡터 이미지를 PNG로 변환
    
    Args:
        data: 원본 벡터 이미지 데이터
        format: 'wmf' 또는 'emf'
        dpi: 출력 해상도
        
    Returns:
        PNG 이미지 데이터 또는 None (변환 실패 시)
    """
    if not HAS_PIL:
        return None
    
    # 방법 1: PIL/Pillow로 직접 변환 시도 (Windows에서만 작동)
    try:
        img = Image.open(io.BytesIO(data))
        output = io.BytesIO()
        img.save(output, format='PNG')
        return output.getvalue()
    except Exception:
        pass
    
    # 방법 2: ImageMagick 사용 (설치되어 있는 경우)
    try:
        result = _convert_with_imagemagick(data, format, dpi)
        if result:
            return result
    except Exception:
        pass
    
    # 방법 3: LibreOffice 사용 (설치되어 있는 경우)
    try:
        result = _convert_with_libreoffice(data, format)
        if result:
            return result
    except Exception:
        pass
    
    # 변환 실패 - 원본 반환
    return None


def _convert_with_imagemagick(data: bytes, format: str, dpi: int = 300) -> Optional[bytes]:
    """ImageMagick을 사용한 변환"""
    try:
        # ImageMagick이 설치되어 있는지 확인
        result = subprocess.run(['which', 'convert'], capture_output=True)
        if result.returncode != 0:
            return None
        
        # 임시 파일로 저장 후 변환
        with tempfile.NamedTemporaryFile(suffix=f'.{format}', delete=False) as tmp_in:
            tmp_in.write(data)
            tmp_in_path = tmp_in.name
        
        tmp_out_path = tmp_in_path.replace(f'.{format}', '.png')
        
        try:
            subprocess.run([
                'convert',
                '-density', str(dpi),
                tmp_in_path,
                '-background', 'white',
                '-flatten',
                tmp_out_path
            ], check=True, capture_output=True)
            
            with open(tmp_out_path, 'rb') as f:
                return f.read()
        finally:
            Path(tmp_in_path).unlink(missing_ok=True)
            Path(tmp_out_path).unlink(missing_ok=True)
            
    except Exception:
        return None


def _convert_with_libreoffice(data: bytes, format: str) -> Optional[bytes]:
    """LibreOffice를 사용한 변환"""
    try:
        # LibreOffice가 설치되어 있는지 확인
        lo_path = None
        for path in ['/Applications/LibreOffice.app/Contents/MacOS/soffice',
                     '/usr/bin/libreoffice', '/usr/bin/soffice']:
            if Path(path).exists():
                lo_path = path
                break
        
        if not lo_path:
            return None
        
        # 임시 디렉토리에서 변환
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_in = Path(tmp_dir) / f'input.{format}'
            tmp_in.write_bytes(data)
            
            subprocess.run([
                lo_path,
                '--headless',
                '--convert-to', 'png',
                '--outdir', tmp_dir,
                str(tmp_in)
            ], check=True, capture_output=True)
            
            tmp_out = Path(tmp_dir) / 'input.png'
            if tmp_out.exists():
                return tmp_out.read_bytes()
                
    except Exception:
        return None
    
    return None


# =============================================================================
# 이미지 형식 감지
# =============================================================================

def _detect_image_format(data: bytes) -> str:
    """이미지 형식 감지"""
    if len(data) < 4:
        return "unknown"
    
    if data[:2] == b'\xff\xd8':
        return "jpg"
    elif data[:4] == b'\x89PNG':
        return "png"
    elif data[:2] == b'BM':
        return "bmp"
    elif data[:4] == b'GIF8':
        return "gif"
    elif data[:4] == b'\xd7\xcd\xc6\x9a':
        return "wmf"
    elif len(data) >= 44 and data[40:44] == b' EMF':
        return "emf"
    elif data[:4] == b'\x01\x00\x00\x00':
        # EMF 시그니처 (다른 형태)
        return "emf"
    else:
        return "unknown"


def _decompress_if_needed(data: bytes, filename: str) -> tuple[bytes, bool]:
    """필요시 zlib 압축 해제"""
    # 먼저 이미지 형식 확인
    fmt = _detect_image_format(data)
    
    if fmt != "unknown":
        return data, False
    
    # 압축 해제 시도
    try:
        decompressed = zlib.decompress(data, -15)
        return decompressed, True
    except zlib.error:
        return data, False


# =============================================================================
# HWP 이미지 추출
# =============================================================================

def extract_images_from_hwp(hwp_file: str | Path) -> list[EmbeddedImage]:
    """
    HWP 파일에서 모든 이미지 추출
    
    Args:
        hwp_file: HWP 파일 경로
        
    Returns:
        list[EmbeddedImage]: 추출된 이미지 목록
    """
    if not HAS_OLEFILE:
        raise ImportError(
            "olefile 라이브러리가 필요합니다.\n"
            "설치: pip install olefile"
        )
    
    hwp_file = Path(hwp_file)
    if not hwp_file.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {hwp_file}")
    
    images = []
    gso_info_list = []  # GSO 정보를 먼저 수집
    
    try:
        ole = olefile.OleFileIO(str(hwp_file))
        
        # 압축 여부 확인
        is_compressed = False
        if ole.exists('FileHeader'):
            header = ole.openstream('FileHeader').read()
            if len(header) >= 36:
                flags = struct.unpack('<I', header[32:36])[0]
                is_compressed = bool(flags & 0x01)
        
        # 1. 먼저 BodyText에서 GSO 정보 수집
        gso_info_list = _collect_gso_info(ole, is_compressed)
        
        # 2. BinData 스트림에서 이미지 추출
        for entry in ole.listdir():
            path = '/'.join(entry)
            
            if not path.startswith('BinData/'):
                continue
            
            # 이미지 데이터 읽기
            raw_data = ole.openstream(entry).read()
            
            # 파일명 추출
            filename = entry[-1]  # BIN0001.jpg
            bin_id = filename.split('.')[0]  # BIN0001
            
            # 압축 해제 (필요시)
            data, was_compressed = _decompress_if_needed(raw_data, filename)
            
            # 이미지 형식 감지
            fmt = _detect_image_format(data)
            
            # 형식이 unknown이면 확장자에서 추측
            if fmt == "unknown":
                ext = filename.split('.')[-1].lower()
                if ext in ['jpg', 'jpeg', 'png', 'bmp', 'gif', 'wmf', 'emf']:
                    fmt = ext
            
            # 올바른 확장자로 파일명 수정
            if fmt != "unknown":
                correct_filename = f"{bin_id}.{fmt}"
            else:
                correct_filename = filename
            
            # 이미지 객체 생성
            img = EmbeddedImage(
                bin_id=bin_id,
                filename=correct_filename,
                format=fmt,
                data=data,
                compressed=was_compressed,
                original_size=len(raw_data),
                element_type="image",
            )
            
            # PIL로 이미지 크기 확인
            if HAS_PIL:
                img.pixel_width, img.pixel_height = img.get_size_from_data()
            
            # GSO 정보와 매핑 (BIN ID 기반)
            bin_num = int(bin_id.replace('BIN', '')) if bin_id.startswith('BIN') else -1
            for gso in gso_info_list:
                if gso.get('bin_index') == bin_num - 1:  # 0-indexed
                    img.x = gso.get('x', 0)
                    img.y = gso.get('y', 0)
                    img.width = gso.get('width', 0)
                    img.height = gso.get('height', 0)
                    img.page = gso.get('page', 0)
                    img.gso_type = gso.get('gso_type', '')
                    img.z_order = gso.get('z_order', 0)
                    break
            
            images.append(img)
        
    finally:
        ole.close()
    
    return images


def _collect_gso_info(ole, is_compressed: bool) -> list[dict]:
    """
    BodyText에서 모든 GSO 컨트롤 정보 수집
    """
    gso_list = []
    section_idx = 0
    
    while ole.exists(f'BodyText/Section{section_idx}'):
        try:
            data = ole.openstream(f'BodyText/Section{section_idx}').read()
            
            if is_compressed and data:
                try:
                    data = zlib.decompress(data, -15)
                except zlib.error:
                    pass
            
            # 섹션에서 GSO 정보 추출
            section_gso = _parse_section_gso_detailed(data, section_idx)
            gso_list.extend(section_gso)
            
        except Exception as e:
            print(f"섹션 파싱 오류: {e}")
        
        section_idx += 1
    
    return gso_list


def _parse_section_gso_detailed(data: bytes, section_idx: int) -> list[dict]:
    """섹션에서 GSO 컨트롤 상세 정보 파싱"""
    gso_list = []
    offset = 0
    gso_index = 0
    current_y_offset = 0.0
    
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
        
        # CTRL_HEADER (0x47)
        if tag_id == 0x47 and size >= 4:
            ctrl_id = record_data[:4].decode('latin-1')
            
            # " osg" = gso (Graphic Shape Object)
            if 'osg' in ctrl_id and size >= 24:
                try:
                    # 기본 위치/크기 정보
                    x = struct.unpack('<i', record_data[8:12])[0]
                    y = struct.unpack('<i', record_data[12:16])[0]
                    w = struct.unpack('<I', record_data[16:20])[0]
                    h = struct.unpack('<I', record_data[20:24])[0]
                    
                    # HWPUNIT to mm 변환
                    x_mm = x * HWPUNIT_TO_MM
                    y_mm = y * HWPUNIT_TO_MM
                    w_mm = w * HWPUNIT_TO_MM
                    h_mm = h * HWPUNIT_TO_MM
                    
                    # GSO 타입 추출 (그림, 사각형, 선 등)
                    gso_type = "picture"  # 기본값
                    if size >= 60:
                        type_code = struct.unpack('<I', record_data[54:58])[0] if size >= 58 else 0
                        if type_code == 0:
                            gso_type = "line"
                        elif type_code == 1:
                            gso_type = "rectangle"
                        elif type_code == 2:
                            gso_type = "ellipse"
                        elif type_code == 3:
                            gso_type = "arc"
                        else:
                            gso_type = "picture"
                    
                    gso_info = {
                        'bin_index': gso_index,
                        'x': max(0, x_mm),  # 음수 좌표 보정
                        'y': max(0, y_mm),
                        'width': w_mm,
                        'height': h_mm,
                        'page': section_idx,
                        'gso_type': gso_type,
                        'z_order': gso_index,
                        'level': level,
                    }
                    gso_list.append(gso_info)
                    gso_index += 1
                    
                except struct.error:
                    pass
    
    return gso_list


# =============================================================================
# HWPX 이미지 추출
# =============================================================================

def extract_images_from_hwpx(hwpx_file: str | Path) -> list[EmbeddedImage]:
    """
    HWPX 파일에서 모든 이미지 추출
    
    Args:
        hwpx_file: HWPX 파일 경로
        
    Returns:
        list[EmbeddedImage]: 추출된 이미지 목록
    """
    hwpx_file = Path(hwpx_file)
    if not hwpx_file.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {hwpx_file}")
    
    images = []
    
    # 임시 디렉토리에 압축 해제
    temp_dir = tempfile.mkdtemp(prefix="hwpx_img_")
    
    try:
        with zipfile.ZipFile(hwpx_file, 'r') as zf:
            zf.extractall(temp_dir)
        
        temp_path = Path(temp_dir)
        
        # BinData 폴더 확인
        bindata_dir = temp_path / "BinData"
        if bindata_dir.exists():
            # 이미지 파일 순회
            for img_file in sorted(bindata_dir.iterdir()):
                if img_file.is_file():
                    data = img_file.read_bytes()
                    
                    fmt = _detect_image_format(data)
                    if fmt == "unknown":
                        fmt = img_file.suffix.lstrip('.').lower()
                    
                    img = EmbeddedImage(
                        bin_id=img_file.stem,
                        filename=img_file.name,
                        format=fmt,
                        data=data,
                        element_type="image",
                    )
                    
                    if HAS_PIL:
                        img.pixel_width, img.pixel_height = img.get_size_from_data()
                    
                    images.append(img)
        
        # Preview 이미지도 포함 (선택적)
        preview_dir = temp_path / "Preview"
        if preview_dir.exists():
            for img_file in preview_dir.glob("*.png"):
                data = img_file.read_bytes()
                
                img = EmbeddedImage(
                    bin_id="Preview",
                    filename=f"preview_{img_file.name}",
                    format="png",
                    data=data,
                    element_type="preview",
                )
                
                if HAS_PIL:
                    img.pixel_width, img.pixel_height = img.get_size_from_data()
                
                images.append(img)
        
        # section*.xml에서 이미지 위치 정보 파싱
        _parse_hwpx_image_positions(temp_path, images)
        
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    return images


def _parse_hwpx_image_positions(temp_path: Path, images: list[EmbeddedImage]):
    """HWPX section*.xml에서 이미지 위치 정보 파싱"""
    contents_dir = temp_path / "Contents"
    if not contents_dir.exists():
        return
    
    # 이미지 ID -> EmbeddedImage 매핑
    image_map = {img.bin_id: img for img in images}
    image_map.update({img.filename: img for img in images})
    
    for section_file in sorted(contents_dir.glob("section*.xml")):
        try:
            tree = ET.parse(section_file)
            root = tree.getroot()
            
            # 섹션 인덱스
            section_idx = int(section_file.stem.replace('section', ''))
            
            # <pic> 태그 찾기
            for elem in root.iter():
                tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                
                if tag == 'pic':
                    _parse_pic_element(elem, image_map, section_idx)
                    
        except ET.ParseError as e:
            print(f"XML 파싱 오류 ({section_file}): {e}")


def _parse_pic_element(pic_elem, image_map: dict, section_idx: int):
    """<hp:pic> 요소에서 이미지 정보 추출"""
    binary_ref = None
    x, y, w, h = 0.0, 0.0, 0.0, 0.0
    
    for child in pic_elem.iter():
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        
        # 바이너리 참조
        if tag == 'imgData':
            binary_ref = child.get('binary', '')
        
        # 오프셋
        elif tag == 'offset':
            x = float(child.get('x', 0)) * HWPUNIT_TO_MM
            y = float(child.get('y', 0)) * HWPUNIT_TO_MM
        
        # 현재 크기
        elif tag == 'curSz':
            w = float(child.get('width', 0)) * HWPUNIT_TO_MM
            h = float(child.get('height', 0)) * HWPUNIT_TO_MM
    
    # 이미지 객체에 위치 정보 업데이트
    if binary_ref and binary_ref in image_map:
        img = image_map[binary_ref]
        img.x = max(0, x)  # 음수 보정
        img.y = max(0, y)
        img.width = w
        img.height = h
        img.page = section_idx


# =============================================================================
# OCR 연동용 출력
# =============================================================================

def save_images_for_ocr(
    images: list[EmbeddedImage],
    output_dir: str | Path,
    json_filename: str = "images_metadata.json",
    convert_vector: bool = True,
) -> tuple[Path, list[Path]]:
    """
    이미지를 OCR 연동용으로 저장
    
    Args:
        images: 추출된 이미지 목록
        output_dir: 출력 디렉토리
        json_filename: 메타데이터 JSON 파일명
        convert_vector: WMF/EMF를 PNG로 변환할지 여부
        
    Returns:
        (JSON 파일 경로, 저장된 이미지 파일 경로 목록)
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    saved_paths = []
    ocr_metadata = {
        "image_count": len(images),
        "images": [],
    }
    
    for img in images:
        # 이미지 저장
        saved_path = img.save(output_dir, convert_vector=convert_vector)
        saved_paths.append(saved_path)
        
        # OCR 메타데이터 추가
        ocr_dict = img.to_ocr_dict()
        ocr_dict["saved_path"] = str(saved_path)
        ocr_metadata["images"].append(ocr_dict)
    
    # JSON 저장
    json_path = output_dir / json_filename
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(ocr_metadata, f, ensure_ascii=False, indent=2)
    
    return json_path, saved_paths


def generate_images_report(images: list[EmbeddedImage], title: str = "") -> str:
    """
    이미지 목록 보고서 생성 (마크다운 형식)
    """
    lines = []
    lines.append(f"# {title or '이미지 추출 결과'}")
    lines.append("")
    lines.append(f"**총 {len(images)}개 이미지**")
    lines.append("")
    
    if not images:
        lines.append("이미지 없음")
        return "\n".join(lines)
    
    # 요약 테이블
    lines.append("| # | 파일명 | 형식 | 크기 | 해상도 | 위치 (mm) | 페이지 |")
    lines.append("|---|--------|------|------|--------|-----------|--------|")
    
    for i, img in enumerate(images, 1):
        size_str = f"{len(img.data):,} B"
        res_str = f"{img.pixel_width}×{img.pixel_height}" if img.pixel_width else "-"
        pos_str = f"({img.x:.1f}, {img.y:.1f})" if img.width > 0 else "-"
        page_str = str(img.page + 1) if img.width > 0 else "-"
        
        lines.append(f"| {i} | {img.filename} | {img.format.upper()} | {size_str} | {res_str} | {pos_str} | {page_str} |")
    
    lines.append("")
    lines.append("## 상세 정보")
    lines.append("")
    
    for i, img in enumerate(images, 1):
        lines.append(f"### {i}. {img.filename}")
        lines.append("")
        lines.append(f"- **형식**: {img.format.upper()}")
        lines.append(f"- **파일 크기**: {len(img.data):,} bytes")
        if img.pixel_width and img.pixel_height:
            lines.append(f"- **해상도**: {img.pixel_width}×{img.pixel_height} px")
        if img.width > 0:
            lines.append(f"- **문서 내 위치**: ({img.x:.2f}, {img.y:.2f}) mm")
            lines.append(f"- **문서 내 크기**: {img.width:.2f}×{img.height:.2f} mm")
            lines.append(f"- **페이지**: {img.page + 1}")
        if img.gso_type:
            lines.append(f"- **GSO 타입**: {img.gso_type}")
        lines.append("")
    
    return "\n".join(lines)


# =============================================================================
# 테스트 및 메인
# =============================================================================

def print_image_summary(images: list[EmbeddedImage], title: str = ""):
    """이미지 목록 요약 출력"""
    print(f"\n{'=' * 60}")
    print(f"📷 {title or '이미지 추출 결과'}")
    print(f"{'=' * 60}")
    
    if not images:
        print("  이미지 없음")
        return
    
    print(f"  총 {len(images)}개 이미지 발견\n")
    
    for i, img in enumerate(images, 1):
        print(f"  [{i}] {img.filename}")
        print(f"      형식: {img.format.upper()}")
        print(f"      크기: {len(img.data):,} bytes")
        if img.pixel_width and img.pixel_height:
            print(f"      해상도: {img.pixel_width}×{img.pixel_height} px")
        if img.width > 0:
            print(f"      문서 내 위치: ({img.x:.1f}, {img.y:.1f}) mm")
            print(f"      문서 내 크기: {img.width:.1f}×{img.height:.1f} mm")
            print(f"      페이지: {img.page + 1}")
        if img.gso_type:
            print(f"      GSO 타입: {img.gso_type}")
        print()


if __name__ == "__main__":
    from pathlib import Path
    
    data_dir = Path(__file__).parent.parent / "data"
    
    # HWP 테스트
    hwp_file = data_dir / "2. [농협] 광고안(B).hwp"
    if hwp_file.exists():
        print(f"\n🔍 HWP 파일 처리: {hwp_file.name}")
        try:
            images = extract_images_from_hwp(hwp_file)
            print_image_summary(images, f"HWP: {hwp_file.name}")
            
            # 이미지 저장 (OCR 연동용)
            if images:
                output_dir = Path(__file__).parent / "results" / "hwp" / "images"
                json_path, saved_paths = save_images_for_ocr(
                    images, output_dir,
                    json_filename=f"{hwp_file.stem}_images.json"
                )
                print(f"  📄 OCR 메타데이터: {json_path}")
                for p in saved_paths:
                    print(f"  💾 저장됨: {p}")
                
                # 마크다운 보고서 저장
                report = generate_images_report(images, f"HWP: {hwp_file.name}")
                report_path = output_dir / f"{hwp_file.stem}_images.md"
                report_path.write_text(report, encoding='utf-8')
                print(f"  📝 보고서: {report_path}")
                    
        except Exception as e:
            import traceback
            print(f"  ❌ 오류: {e}")
            traceback.print_exc()
    
    # HWPX 테스트
    hwpx_file = data_dir / "은행권 광고심의 결과 보고서(양식)vF (1).hwpx"
    if hwpx_file.exists():
        print(f"\n🔍 HWPX 파일 처리: {hwpx_file.name}")
        try:
            images = extract_images_from_hwpx(hwpx_file)
            print_image_summary(images, f"HWPX: {hwpx_file.name}")
            
            # 이미지 저장 (OCR 연동용)
            if images:
                output_dir = Path(__file__).parent / "results" / "hwpx" / "images"
                json_path, saved_paths = save_images_for_ocr(
                    images, output_dir,
                    json_filename=f"{hwpx_file.stem}_images.json"
                )
                print(f"  📄 OCR 메타데이터: {json_path}")
                for p in saved_paths:
                    print(f"  💾 저장됨: {p}")
                
                # 마크다운 보고서 저장
                report = generate_images_report(images, f"HWPX: {hwpx_file.name}")
                report_path = output_dir / f"{hwpx_file.stem}_images.md"
                report_path.write_text(report, encoding='utf-8')
                print(f"  📝 보고서: {report_path}")
                    
        except Exception as e:
            import traceback
            print(f"  ❌ 오류: {e}")
            traceback.print_exc()
    
    print("\n✅ 완료!")
