import os
import re
import time
from typing import List, Dict, Any, Tuple
from docx import Document
from docx.oxml.shared import OxmlElement, qn
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from config import Config
import hashlib
import subprocess
import shutil
import copy

class DocumentProcessor:
    def __init__(self):
        self.highlight_color = Config.HIGHLIGHT_COLOR

    def extract_document_content(self, file_path: str) -> str:
        doc = Document(file_path)
        content = []
        for table in doc.tables:
            for row in table.rows:
                row_content = []
                for cell in row.cells:
                    row_content.append(cell.text.strip())
                content.append(' | '.join(row_content))

        print('\n'.join(content))
        return '\n'.join(content)

    def _save_cell_format(self, cell):
        """Saving cell formatting information"""
        format_info = {
            'paragraphs': []
        }
        
        for paragraph in cell.paragraphs:
            para_format = {
                'alignment': paragraph.alignment,
                'runs': []
            }
            
            for run in paragraph.runs:
                run_format = {
                    'text': run.text,
                    'bold': run.bold,
                    'italic': run.italic,
                    'underline': run.underline,
                    'font_name': run.font.name,
                    'font_size': run.font.size,
                    'font_color': run.font.color.rgb if run.font.color.rgb else None
                }
                para_format['runs'].append(run_format)
            
            format_info['paragraphs'].append(para_format)
        
        return format_info

    def _apply_cell_format(self, cell, format_info, new_text):
        """Apply saved formatting to cells"""
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.text = ""
        
        if not cell.paragraphs:
            cell.add_paragraph()
        
        if format_info['paragraphs']:
            para_format = format_info['paragraphs'][0]
            paragraph = cell.paragraphs[0]
            paragraph.alignment = para_format['alignment']
            
            if para_format['runs']:
                run_format = para_format['runs'][0]
                run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
                run.text = new_text
                run.bold = run_format['bold']
                run.italic = run_format['italic']
                run.underline = run_format['underline']
                if run_format['font_name']:
                    run.font.name = run_format['font_name']
                if run_format['font_size']:
                    run.font.size = run_format['font_size']
                if run_format['font_color']:
                    run.font.color.rgb = run_format['font_color']
            else:
                paragraph.text = new_text
        else:
            cell.paragraphs[0].text = new_text

    def find_and_number_all_fields_docx(self, file_path: str) -> Tuple[List[Dict[str, Any]], str]:
        """Adding index to all cells in the document, merged cells share the same index (based on whether cell.text already has a number to determine)."""
        doc = Document(file_path)
        all_fields = []
        field_index = 1
        key_to_index = {}

        for table_index, table in enumerate(doc.tables):
            for row_index, row in enumerate(table.rows):
                for col_index, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    if re.search(r'\[\d+\]$', cell_text):
                        continue
                    
                    original_format = self._save_cell_format(cell)
                    
                    key = id(cell._tc)
                    key_to_index[key] = field_index
                    
                    if cell_text:
                        new_text = f"{cell_text} [{field_index}]"
                    else:
                        new_text = f"[{field_index}]"
                    
                    self._apply_cell_format(cell, original_format, new_text)
                    
                    all_fields.append({
                        'index': field_index,
                        'type': 'table_cell',
                        'table_index': table_index,
                        'row_index': row_index,
                        'col_index': col_index,
                        'text': cell_text,
                        'original_format': original_format,
                        'context': f'Table {table_index+1}, Row {row_index+1}, Col {col_index+1}'
                    })
                    field_index += 1

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(Config.MID_DIR, f"{base_name}_numbered_{int(time.time())}.docx")
        doc.save(output_path)
        return all_fields, output_path

    def _create_shading_element(self):
        """Yellow background shading element for Word cell."""
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), 'FFFF00')
        return shading

    def fill_document(self, file_path: str, field_answers: List[Dict[str, Any]]) -> str:
        """Fill document using field_answers that include table/row/col index."""
        from docx import Document
        import os
        from config import Config
        doc = Document(file_path)

        for ans in field_answers:
            try:
                table_idx = ans["table_index"]
                row_idx = ans["row_index"]
                col_idx = ans["col_index"]
                content = ans["content"]
                cell = doc.tables[table_idx].cell(row_idx, col_idx)
                
                original_format = ans.get("original_format")
                if original_format:
                    self._apply_cell_format(cell, original_format, str(content))
                else:
                    cell.text = str(content)
            except Exception as e:
                print(f"Failed to fill cell index {ans.get('index')}: {e}")

        original_filename = os.path.basename(file_path)
        name_without_ext = original_filename.split("_numbered")[0].split("_highlighted")[0].split("_restored")[0].replace(".docx", "")
        timestamp = int(time.time())
        filled_filename = f"{name_without_ext}_filled_{timestamp}.docx"
        filled_path = os.path.join(Config.OUTPUT_DIR, filled_filename)
        doc.save(filled_path)
        return filled_path

    def restore_cells_content_from_indexed_docx(self, file_path: str, restored_cells: List[Dict[str, Any]]) -> str:
        """Restore cell content for cells that do not need to be filled, based on their index."""
        doc = Document(file_path)
        for cell_info in restored_cells:
            idx = cell_info.get("index")
            content = cell_info.get("restored_content", "")
            original_format = cell_info.get("original_format")
            found = False
            for table_index, table in enumerate(doc.tables):
                for row_index, row in enumerate(table.rows):
                    for col_index, cell in enumerate(row.cells):
                        cell_text = cell.text.strip()
                        if cell_text.endswith(f"[{idx}]") or cell_text == f"[{idx}]":
                            if original_format:
                                self._apply_cell_format(cell, original_format, content)
                            else:
                                cell.text = content
                            found = True
                            break
                    if found:
                        break
                if found:
                    break
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(Config.MID_DIR, f"{base_name}_restored_{int(time.time())}.docx")
        doc.save(output_path)
        return output_path

    def convert_doc_to_docx(self, doc_path, docx_path):
        """Convert doc to docx using libreoffice"""
        output_dir = os.path.dirname(docx_path)
        subprocess.run(['C:\Program Files\LibreOffice\program\soffice.exe', '--headless', '--convert-to', 'docx', doc_path, '--outdir', output_dir], check=True)
        base = os.path.splitext(os.path.basename(doc_path))[0]
        generated_docx = os.path.join(output_dir, base + '.docx')
        if generated_docx != docx_path:
            shutil.move(generated_docx, docx_path)
        return docx_path

