import os
import re
import time
from typing import List, Dict, Any, Tuple
from docx import Document
from docx.oxml.shared import OxmlElement, qn
from config import Config
import hashlib

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

    def find_and_number_all_fields_docx(self, file_path: str) -> Tuple[List[Dict[str, Any]], str]:
        """Adding index to all cells in the document, merged cells share the same index (using cell._tc.xml hash as unique id, stable order)."""
        doc = Document(file_path)
        all_fields = []
        field_index = 1
        cell_hash_to_index = {}

        for table_index, table in enumerate(doc.tables):
            for row_index, row in enumerate(table.rows):
                for col_index, cell in enumerate(row.cells):
                    cell_hash = hashlib.md5(cell._tc.xml.encode('utf-8')).hexdigest()
                    cell_text = cell.text.strip()
                    if cell_hash not in cell_hash_to_index:
                        cell_hash_to_index[cell_hash] = field_index
                        # Only write index for the first occurrence
                        if cell_text:
                            cell.text = f"{cell_text} [{field_index}]"
                        else:
                            cell.text = f"[{field_index}]"
                        all_fields.append({
                            'index': field_index,
                            'type': 'table_cell',
                            'table_index': table_index,
                            'row_index': row_index,
                            'col_index': col_index,
                            'text': cell_text,
                            'context': f'Table {table_index+1}, Row {row_index+1}, Col {col_index+1}'
                        })
                        field_index += 1
                    else:
                        pass

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(Config.MID_DIR, f"{base_name}_numbered_{int(time.time())}.docx")
        doc.save(output_path)
        return all_fields, output_path

    def highlight_empty_fields_docx(self, file_path: str, empty_fields: List[Dict[str, Any]]) -> str:
        doc = Document(file_path)
        for field in empty_fields:
            if field['type'] == 'table_cell':
                table = doc.tables[field['table_index']]
                cell = table.cell(field['row_index'], field['col_index'])
                cell._tc.get_or_add_tcPr().append(self._create_shading_element())
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(Config.MID_DIR, f"{base_name}_highlighted_{int(time.time())}.docx")
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
            found = False
            for table_index, table in enumerate(doc.tables):
                for row_index, row in enumerate(table.rows):
                    for col_index, cell in enumerate(row.cells):
                        cell_text = cell.text.strip()
                        if cell_text.endswith(f"[{idx}]") or cell_text == f"[{idx}]":
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

