"""
HWPX Parser 다이어그램 생성 스크립트

이 스크립트는 hwpx_folder_parser_cursor.py의 구조를 시각화합니다.
graphviz가 설치되어 있어야 합니다.

설치: pip install graphviz
시스템 설치: brew install graphviz (macOS)
"""

from graphviz import Digraph
from pathlib import Path


def create_class_diagram():
    """클래스 다이어그램 생성"""
    dot = Digraph(comment='HWPX Parser Class Diagram')
    dot.attr(rankdir='TB', splines='ortho')
    dot.attr('node', shape='record', fontname='Helvetica', fontsize='10')
    dot.attr('edge', fontname='Helvetica', fontsize='9')
    
    # 레이아웃 클래스들 (노란색 계열)
    with dot.subgraph(name='cluster_layout') as c:
        c.attr(label='레이아웃 클래스', style='filled', color='#FFF8E1', fontname='Helvetica-Bold')
        
        c.node('Position', '''Position|
vert_rel_to: str\\l
horz_rel_to: str\\l
vert_align: str\\l
horz_align: str\\l
vert_offset: int\\l
horz_offset: int\\l|
+ to_mm(): dict\\l''')
        
        c.node('Size', '''Size|
width: int\\l
height: int\\l
width_rel_to: str\\l
height_rel_to: str\\l|
+ to_mm(): dict\\l''')
        
        c.node('Margin', '''Margin|
left: int\\l
right: int\\l
top: int\\l
bottom: int\\l|
+ to_mm(): dict\\l''')
        
        c.node('LineSegment', '''LineSegment|
text_pos: int\\l
vert_pos: int\\l
horz_pos: int\\l
vert_size: int\\l
horz_size: int\\l|
+ to_mm(): dict\\l''')
        
        c.node('PageProperties', '''PageProperties|
width: int\\l
height: int\\l
landscape: str\\l
margin: Margin\\l|
+ to_mm(): dict\\l''')
    
    # 콘텐츠 클래스들 (녹색 계열)
    with dot.subgraph(name='cluster_content') as c:
        c.attr(label='콘텐츠 클래스', style='filled', color='#E8F5E9', fontname='Helvetica-Bold')
        
        c.node('TableCell', '''TableCell|
row: int\\l
col: int\\l
text: str\\l
size: Size\\l
margin: Margin\\l''')
        
        c.node('Table', '''Table|
rows: int\\l
cols: int\\l
cells: list[TableCell]\\l
position: Position\\l
size: Size\\l|
+ to_markdown()\\l
+ to_markdown_with_layout()\\l''')
        
        c.node('TextRun', '''TextRun|
text: str\\l
char_pr_id: str\\l''')
        
        c.node('Paragraph', '''Paragraph|
id: str\\l
texts: list[str]\\l
text_runs: list[TextRun]\\l
tables: list[Table]\\l
line_segments: list[LineSegment]\\l|
+ full_text: str\\l
+ get_bounding_box(): dict\\l''')
        
        c.node('Section', '''Section|
index: int\\l
paragraphs: list[Paragraph]\\l
page_props: PageProperties\\l|
+ full_text: str\\l''')
    
    # 문서 클래스 (파란색 계열)
    with dot.subgraph(name='cluster_document') as c:
        c.attr(label='문서 클래스', style='filled', color='#E3F2FD', fontname='Helvetica-Bold')
        
        c.node('VersionInfo', '''VersionInfo|
application: str\\l
app_version: str\\l
xml_version: str\\l''')
        
        c.node('HwpxDocument', '''HwpxDocument|
folder_path: Path\\l
version: VersionInfo\\l
sections: list[Section]\\l
metadata: dict\\l|
+ to_text(): str\\l
+ to_markdown(): str\\l
+ to_markdown_with_layout(): str\\l
+ to_json(): str\\l
+ to_json_with_layout(): str\\l''')
    
    # 파서 클래스 (주황색 계열)
    with dot.subgraph(name='cluster_parser') as c:
        c.attr(label='파서 클래스', style='filled', color='#FFF3E0', fontname='Helvetica-Bold')
        
        c.node('HwpxFolderParser', '''HwpxFolderParser|
folder_path: Path\\l
contents_dir: Path\\l|
+ parse(): HwpxDocument\\l
- _parse_version()\\l
- _parse_metadata()\\l
- _parse_sections()\\l
- _parse_section()\\l
- _parse_paragraph()\\l
- _parse_table()\\l
- _parse_table_cell()\\l''')
    
    # 관계 정의
    dot.edge('HwpxFolderParser', 'HwpxDocument', label='creates', style='dashed')
    dot.edge('HwpxDocument', 'VersionInfo', label='has')
    dot.edge('HwpxDocument', 'Section', label='has many')
    dot.edge('Section', 'PageProperties', label='has')
    dot.edge('Section', 'Paragraph', label='has many')
    dot.edge('PageProperties', 'Margin', label='has')
    dot.edge('Paragraph', 'TextRun', label='has many')
    dot.edge('Paragraph', 'Table', label='has many')
    dot.edge('Paragraph', 'LineSegment', label='has many')
    dot.edge('Table', 'TableCell', label='has many')
    dot.edge('Table', 'Position', label='has')
    dot.edge('Table', 'Size', label='has')
    dot.edge('TableCell', 'Size', label='has')
    dot.edge('TableCell', 'Margin', label='has')
    
    return dot


def create_flow_diagram():
    """파싱 흐름도 생성"""
    dot = Digraph(comment='HWPX Parsing Flow')
    dot.attr(rankdir='TB')
    dot.attr('node', fontname='Helvetica', fontsize='10')
    dot.attr('edge', fontname='Helvetica', fontsize='9')
    
    # 입력 파일들
    with dot.subgraph(name='cluster_input') as c:
        c.attr(label='📁 HWPX 폴더', style='filled', color='#E1F5FE', fontname='Helvetica-Bold')
        c.node('version_xml', 'version.xml', shape='note')
        c.node('header_xml', 'header.xml', shape='note')
        c.node('section_xml', 'section0.xml', shape='note')
    
    # 파싱 단계
    with dot.subgraph(name='cluster_parse') as c:
        c.attr(label='🔧 파싱 단계', style='filled', color='#FFF3E0', fontname='Helvetica-Bold')
        c.node('parse', 'parse()', shape='box', style='rounded')
        c.node('parse_version', '_parse_version()', shape='box', style='rounded')
        c.node('parse_metadata', '_parse_metadata()', shape='box', style='rounded')
        c.node('parse_sections', '_parse_sections()', shape='box', style='rounded')
        c.node('parse_section', '_parse_section()', shape='box', style='rounded')
        c.node('parse_para', '_parse_paragraph()', shape='box', style='rounded')
        c.node('parse_table', '_parse_table()', shape='box', style='rounded')
        c.node('parse_cell', '_parse_table_cell()', shape='box', style='rounded')
    
    # 출력
    with dot.subgraph(name='cluster_output') as c:
        c.attr(label='📄 출력', style='filled', color='#E8F5E9', fontname='Helvetica-Bold')
        c.node('doc', 'HwpxDocument', shape='box3d')
        c.node('to_text', 'to_text()', shape='box', style='rounded')
        c.node('to_md', 'to_markdown()', shape='box', style='rounded')
        c.node('to_md_layout', 'to_markdown_with_layout()', shape='box', style='rounded')
        c.node('to_json', 'to_json()', shape='box', style='rounded')
        c.node('to_json_layout', 'to_json_with_layout()', shape='box', style='rounded')
    
    # 연결
    dot.edge('version_xml', 'parse_version')
    dot.edge('header_xml', 'parse_metadata')
    dot.edge('section_xml', 'parse_sections')
    
    dot.edge('parse', 'parse_version')
    dot.edge('parse', 'parse_metadata')
    dot.edge('parse', 'parse_sections')
    
    dot.edge('parse_sections', 'parse_section')
    dot.edge('parse_section', 'parse_para')
    dot.edge('parse_para', 'parse_table')
    dot.edge('parse_table', 'parse_cell')
    
    dot.edge('parse', 'doc')
    dot.edge('doc', 'to_text')
    dot.edge('doc', 'to_md')
    dot.edge('doc', 'to_md_layout')
    dot.edge('doc', 'to_json')
    dot.edge('doc', 'to_json_layout')
    
    return dot


def create_data_flow_diagram():
    """데이터 흐름도 생성"""
    dot = Digraph(comment='Data Flow')
    dot.attr(rankdir='LR')
    dot.attr('node', fontname='Helvetica', fontsize='10')
    
    # 계층 구조
    dot.node('doc', '📄 HwpxDocument', shape='folder', style='filled', fillcolor='#E3F2FD')
    dot.node('ver', 'VersionInfo', shape='box', style='filled', fillcolor='#E3F2FD')
    dot.node('sec', '📑 Section', shape='folder', style='filled', fillcolor='#E8F5E9')
    dot.node('page', 'PageProperties', shape='box', style='filled', fillcolor='#FFF8E1')
    dot.node('para', '📝 Paragraph', shape='folder', style='filled', fillcolor='#E8F5E9')
    dot.node('lseg', 'LineSegment', shape='box', style='filled', fillcolor='#FFF8E1')
    dot.node('trun', 'TextRun', shape='box', style='filled', fillcolor='#E8F5E9')
    dot.node('tbl', '📊 Table', shape='folder', style='filled', fillcolor='#E8F5E9')
    dot.node('pos', 'Position', shape='box', style='filled', fillcolor='#FFF8E1')
    dot.node('size', 'Size', shape='box', style='filled', fillcolor='#FFF8E1')
    dot.node('cell', '🔲 TableCell', shape='box', style='filled', fillcolor='#E8F5E9')
    dot.node('margin', 'Margin', shape='box', style='filled', fillcolor='#FFF8E1')
    
    dot.edge('doc', 'ver')
    dot.edge('doc', 'sec')
    dot.edge('sec', 'page')
    dot.edge('sec', 'para')
    dot.edge('para', 'lseg')
    dot.edge('para', 'trun')
    dot.edge('para', 'tbl')
    dot.edge('tbl', 'pos')
    dot.edge('tbl', 'size')
    dot.edge('tbl', 'cell')
    dot.edge('cell', 'margin')
    dot.edge('page', 'margin')
    
    return dot


def main():
    """메인 함수"""
    output_dir = Path(__file__).parent / "diagrams"
    output_dir.mkdir(exist_ok=True)
    
    print("📊 다이어그램 생성 중...")
    
    # 클래스 다이어그램
    print("  1. 클래스 다이어그램...")
    class_diagram = create_class_diagram()
    class_diagram.render(output_dir / 'class_diagram', format='png', cleanup=True)
    class_diagram.render(output_dir / 'class_diagram', format='svg', cleanup=True)
    
    # 파싱 흐름도
    print("  2. 파싱 흐름도...")
    flow_diagram = create_flow_diagram()
    flow_diagram.render(output_dir / 'parsing_flow', format='png', cleanup=True)
    flow_diagram.render(output_dir / 'parsing_flow', format='svg', cleanup=True)
    
    # 데이터 흐름도
    print("  3. 데이터 흐름도...")
    data_flow = create_data_flow_diagram()
    data_flow.render(output_dir / 'data_flow', format='png', cleanup=True)
    data_flow.render(output_dir / 'data_flow', format='svg', cleanup=True)
    
    print(f"\n✅ 다이어그램 저장 완료: {output_dir}/")
    print("   - class_diagram.png/svg")
    print("   - parsing_flow.png/svg")
    print("   - data_flow.png/svg")


if __name__ == "__main__":
    try:
        main()
    except ImportError:
        print("❌ graphviz 라이브러리가 설치되어 있지 않습니다.")
        print("   설치 방법:")
        print("   1. pip install graphviz")
        print("   2. brew install graphviz  (macOS)")
        print("   3. apt install graphviz   (Linux)")
        print("\n대신 hwpx_parser_diagram.md 파일을 확인해주세요.")
        print("GitHub나 VS Code의 Mermaid 확장으로 볼 수 있습니다.")

