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
        """Find empty fields and write [index] into them."""
        doc = Document(file_path)
        empty_fields = []
        field_index = 1

        for table_index, table in enumerate(doc.tables):
            for row_index, row in enumerate(table.rows):
                for col_index, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    if not cell_text or cell_text in ['To be filled', 'Blank', '_____', '待填写', '空白']:
                        cell.text = f"[{field_index}]"
                        empty_fields.append({
                            'index': field_index,
                            'type': 'table_cell',
                            'table_index': table_index,
                            'row_index': row_index,
                            'col_index': col_index,
                            'text': cell_text,
                            'context': f'Table {table_index+1}, Row {row_index+1}, Col {col_index+1}'
                        })
                        field_index += 1

        output_path = file_path.replace('.docx', f'_numbered_{int(time.time())}.docx')
        doc.save(output_path)
        return empty_fields, output_path

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

    def fill_document(self, file_path: str, field_answers: List[Dict[str, Any]]) -> str:
        """Fill document using field_answers that include table/row/col index."""
        from docx import Document
        doc = Document(file_path)

        for ans in field_answers:
            try:
                table_idx = ans["table_index"]
                row_idx = ans["row_index"]
                col_idx = ans["col_index"]
                content = ans["content"]

                cell = doc.tables[table_idx].cell(row_idx, col_idx)
                cell.text = str(content)
            except Exception as e:
                print(f"Failed to fill cell index {ans.get('index')}: {e}")

        original_filename = os.path.basename(file_path)
        name_without_ext = original_filename.replace(".docx", "")
        timestamp = int(time.time())
        filled_filename = f"{name_without_ext}_filled_{timestamp}.docx"
        filled_path = os.path.join(Config.OUTPUT_DIR, filled_filename)
        doc.save(filled_path)
        return filled_path

