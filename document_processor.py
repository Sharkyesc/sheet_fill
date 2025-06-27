import os
import re
import time
from typing import List, Dict, Any, Tuple
from docx import Document
from docx.oxml.shared import OxmlElement, qn
from config import Config

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

    def find_and_number_empty_fields_docx(self, file_path: str) -> Tuple[List[Dict[str, Any]], str]:
        """Find all fields and write [index] into them. For non-empty cells, append [index]."""
        doc = Document(file_path)
        all_fields = []
        field_index = 1

        for table_index, table in enumerate(doc.tables):
            for row_index, row in enumerate(table.rows):
                for col_index, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    original_text = cell_text
                    
                    # 为每个单元格添加索引
                    if not cell_text :
                        cell.text = f"[{field_index}]"
                    else:
                        cell.text = f"{cell_text} [{field_index}]"
                    
                    all_fields.append({
                        'index': field_index,
                        'type': 'table_cell',
                        'table_index': table_index,
                        'row_index': row_index,
                        'col_index': col_index,
                        'text': original_text,
                        'context': f'Table {table_index+1}, Row {row_index+1}, Col {col_index+1}',
                        'is_empty': not original_text or original_text in ['To be filled', 'Blank', '_____', '待填写', '空白']
                    })
                    field_index += 1

        output_path = file_path.replace('.docx', f'_numbered_{int(time.time())}.docx')
        doc.save(output_path)
        return all_fields, output_path

    def highlight_empty_fields_docx(self, file_path: str, empty_fields: List[Dict[str, Any]]) -> str:
        doc = Document(file_path)
        for field in empty_fields:
            if field['type'] == 'table_cell':
                table = doc.tables[field['table_index']]
                cell = table.cell(field['row_index'], field['col_index'])
                cell._tc.get_or_add_tcPr().append(self._create_shading_element())
        output_path = file_path.replace('.docx', f'_highlighted_{int(time.time())}.docx')
        doc.save(output_path)
        return output_path

    def _create_shading_element(self):
        """Yellow background shading element for Word cell."""
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), 'FFFF00')
        return shading

    def list_docx_files(self) -> List[str]:
        """List all .docx files from the input directory."""
        input_dir = Config.INPUT_DIR
        if not os.path.exists(input_dir):
            return []
        return [
            os.path.join(input_dir, file)
            for file in os.listdir(input_dir)
            if file.lower().endswith('.docx')
        ]

    def restore_original_content(self, file_path: str, fields_to_restore: List[Dict[str, Any]]) -> str:
        """Restore original content for fields that don't need filling."""
        doc = Document(file_path)
        
        for field in fields_to_restore:
            try:
                table_idx = field["table_index"]
                row_idx = field["row_index"]
                col_idx = field["col_index"]
                original_content = field["original_content"]
                
                cell = doc.tables[table_idx].cell(row_idx, col_idx)
                cell.text = str(original_content)
            except Exception as e:
                print(f"恢复单元格索引 {field.get('index')} 失败: {e}")
        
        output_path = file_path.replace('.docx', f'_restored_{int(time.time())}.docx')
        doc.save(output_path)
        return output_path

    def fill_document(self, file_path: str, field_answers: List[Dict[str, Any]]) -> str:
        """Fill document using field_answers that include table/row/col index."""
        from docx import Document
        doc = Document(file_path)

        for ans in field_answers:
            try:
                table_idx = ans["table_index"]
                row_idx = ans["row_index"]
                col_idx = ans["col_index"]
                
                # 判断是填充新内容还是恢复原内容
                if "content" in ans:
                    # 填充新内容
                    content = ans["content"]
                    cell = doc.tables[table_idx].cell(row_idx, col_idx)
                    cell.text = str(content)
                elif "original_content" in ans:
                    # 恢复原内容
                    original_content = ans["original_content"]
                    cell = doc.tables[table_idx].cell(row_idx, col_idx)
                    cell.text = str(original_content)
            except Exception as e:
                print(f"填充单元格索引 {ans.get('index')} 失败: {e}")

        original_filename = os.path.basename(file_path)
        name_without_ext = original_filename.replace(".docx", "")
        timestamp = int(time.time())
        filled_filename = f"{name_without_ext}_filled_{timestamp}.docx"
        filled_path = os.path.join(Config.OUTPUT_DIR, filled_filename)
        doc.save(filled_path)
        return filled_path

