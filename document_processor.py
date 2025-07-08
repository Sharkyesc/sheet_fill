import os
import re
import time
from typing import List, Dict, Any, Tuple
from docx import Document
from docx.oxml.shared import OxmlElement, qn
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.cell import MergedCell
from config import Config

class DocumentProcessor:
    def __init__(self):
        self.highlight_color = Config.HIGHLIGHT_COLOR

    def extract_document_content(self, file_path: str) -> str:
        """提取文档内容，支持Word和Excel"""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.docx':
            return self._extract_docx_content(file_path)
        elif file_ext in ['.xlsx', '.xls']:
            return self._extract_excel_content(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_ext}")

    def _extract_docx_content(self, file_path: str) -> str:
        """提取Word文档内容"""
        doc = Document(file_path)
        content = []
        for table in doc.tables:
            for row in table.rows:
                row_content = []
                for cell in row.cells:
                    row_content.append(cell.text.strip())
                content.append(' | '.join(row_content))
        return '\n'.join(content)

    def _extract_excel_content(self, file_path: str) -> str:
        """提取Excel文档内容"""
        wb = openpyxl.load_workbook(file_path)
        content = []
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            content.append(f"=== Sheet: {sheet_name} ===")
            
            for row in ws.iter_rows():
                row_content = []
                for cell in row:
                    cell_value = str(cell.value) if cell.value is not None else ""
                    row_content.append(cell_value.strip())
                content.append(' | '.join(row_content))
        
        wb.close()
        return '\n'.join(content)

    def find_and_number_all_fields_docx(self, file_path: str) -> Tuple[List[Dict[str, Any]], str]:
        """为Word文档所有字段添加索引"""
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
                    key = id(cell._tc)
                    key_to_index[key] = field_index
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

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(Config.MID_DIR, f"{base_name}_numbered_{int(time.time())}.docx")
        doc.save(output_path)
        return all_fields, output_path

    def find_and_number_all_fields_excel(self, file_path: str) -> Tuple[List[Dict[str, Any]], str]:
        """为Excel文档所有字段添加索引"""
        wb = openpyxl.load_workbook(file_path)
        all_fields = []
        field_index = 1

        for sheet_index, sheet_name in enumerate(wb.sheetnames):
            ws = wb[sheet_name]
            
            for row_index, row in enumerate(ws.iter_rows(), 1):
                for col_index, cell in enumerate(row, 1):
                    # 跳过合并单元格中的非主单元格
                    if isinstance(cell, MergedCell):
                        print(f"Skipping merged cell at Row {row_index}, Col {col_index}")
                        continue
                    
                    cell_value = str(cell.value) if cell.value is not None else ""
                    
                    # 检查是否已经有索引标记
                    if re.search(r'\[\d+\]$', cell_value):
                        continue
                    
                    # 添加索引标记
                    try:
                        if cell_value.strip():
                            cell.value = f"{cell_value} [{field_index}]"
                        else:
                            cell.value = f"[{field_index}]"
                        
                        all_fields.append({
                            'index': field_index,
                            'type': 'excel_cell',
                            'sheet_index': sheet_index,
                            'sheet_name': sheet_name,
                            'row_index': row_index,
                            'col_index': col_index,
                            'text': cell_value.strip(),
                            'context': f'Sheet {sheet_name}, Row {row_index}, Col {col_index}'
                        })
                        field_index += 1
                        
                    except Exception as e:
                        print(f"Warning: Cannot modify cell at Row {row_index}, Col {col_index}: {e}")
                        continue

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(Config.MID_DIR, f"{base_name}_numbered_{int(time.time())}.xlsx")
        wb.save(output_path)
        wb.close()
        return all_fields, output_path

    def find_and_number_all_fields(self, file_path: str) -> Tuple[List[Dict[str, Any]], str]:
        """统一的字段标记接口，自动识别文件类型"""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.docx':
            return self.find_and_number_all_fields_docx(file_path)
        elif file_ext in ['.xlsx', '.xls']:
            return self.find_and_number_all_fields_excel(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_ext}")

    def highlight_empty_fields_docx(self, file_path: str, empty_fields: List[Dict[str, Any]]) -> str:
        """Word文档字段高亮"""
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

    def highlight_empty_fields_excel(self, file_path: str, empty_fields: List[Dict[str, Any]]) -> str:
        """Excel文档字段高亮"""
        wb = openpyxl.load_workbook(file_path)
        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        
        for field in empty_fields:
            if field['type'] == 'excel_cell':
                try:
                    ws = wb[field['sheet_name']]
                    cell = ws.cell(row=field['row_index'], column=field['col_index'])
                    
                    # 跳过合并单元格
                    if isinstance(cell, MergedCell):
                        print(f"Warning: Cannot highlight merged cell at Row {field['row_index']}, Col {field['col_index']}")
                        continue
                        
                    cell.fill = yellow_fill
                except Exception as e:
                    print(f"Warning: Cannot highlight cell: {e}")
                    continue
                
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(Config.MID_DIR, f"{base_name}_highlighted_{int(time.time())}.xlsx")
        wb.save(output_path)
        wb.close()
        return output_path

    def _create_shading_element(self):
        """Word单元格黄色背景着色元素"""
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), 'FFFF00')
        return shading

    def list_document_files(self) -> List[str]:
        """列出输入目录中的所有文档文件（Word和Excel）"""
        input_dir = Config.INPUT_DIR
        if not os.path.exists(input_dir):
            return []
        
        supported_extensions = ['.docx', '.xlsx', '.xls']
        files = []
        
        for file in os.listdir(input_dir):
            file_ext = os.path.splitext(file)[1].lower()
            if file_ext in supported_extensions:
                files.append(os.path.join(input_dir, file))
        
        return files

    def list_docx_files(self) -> List[str]:
        """保持向后兼容性的方法"""
        return [f for f in self.list_document_files() if f.endswith('.docx')]

    def fill_document(self, file_path: str, field_answers: List[Dict[str, Any]]) -> str:
        """填充文档，支持Word和Excel"""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.docx':
            return self._fill_docx_document(file_path, field_answers)
        elif file_ext in ['.xlsx', '.xls']:
            return self._fill_excel_document(file_path, field_answers)
        else:
            raise ValueError(f"Unsupported file format: {file_ext}")

    def _fill_docx_document(self, file_path: str, field_answers: List[Dict[str, Any]]) -> str:
        """填充Word文档"""
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

    def _fill_excel_document(self, file_path: str, field_answers: List[Dict[str, Any]]) -> str:
        """填充Excel文档"""
        wb = openpyxl.load_workbook(file_path)

        for ans in field_answers:
            try:
                sheet_name = ans["sheet_name"]
                row_idx = ans["row_index"]
                col_idx = ans["col_index"]
                content = ans["content"]
                
                ws = wb[sheet_name]
                cell = ws.cell(row=row_idx, column=col_idx)
                
                # 检查是否是合并单元格
                if isinstance(cell, MergedCell):
                    print(f"Warning: Cannot fill merged cell at Row {row_idx}, Col {col_idx}")
                    continue
                    
                cell.value = str(content)
                
            except Exception as e:
                print(f"Failed to fill cell index {ans.get('index')}: {e}")

        original_filename = os.path.basename(file_path)
        name_without_ext = original_filename.split("_numbered")[0].split("_highlighted")[0].split("_restored")[0]
        name_without_ext = os.path.splitext(name_without_ext)[0]
        timestamp = int(time.time())
        filled_filename = f"{name_without_ext}_filled_{timestamp}.xlsx"
        filled_path = os.path.join(Config.OUTPUT_DIR, filled_filename)
        wb.save(filled_path)
        wb.close()
        return filled_path

    def restore_cells_content_from_indexed_docx(self, file_path: str, restored_cells: List[Dict[str, Any]]) -> str:
        """恢复Word文档单元格内容"""
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

    def restore_cells_content_from_indexed_excel(self, file_path: str, restored_cells: List[Dict[str, Any]]) -> str:
        """恢复Excel文档单元格内容"""
        wb = openpyxl.load_workbook(file_path)
        
        for cell_info in restored_cells:
            idx = cell_info.get("index")
            content = cell_info.get("restored_content", "")
            found = False
            
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for row in ws.iter_rows():
                    for cell in row:
                        # 跳过合并单元格
                        if isinstance(cell, MergedCell):
                            continue
                            
                        cell_value = str(cell.value) if cell.value is not None else ""
                        if cell_value.endswith(f"[{idx}]") or cell_value == f"[{idx}]":
                            try:
                                cell.value = content
                                found = True
                                break
                            except Exception as e:
                                print(f"Warning: Cannot restore cell with index {idx}: {e}")
                                continue
                    if found:
                        break
                if found:
                    break
        
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(Config.MID_DIR, f"{base_name}_restored_{int(time.time())}.xlsx")
        wb.save(output_path)
        wb.close()
        return output_path

    def restore_cells_content_from_indexed(self, file_path: str, restored_cells: List[Dict[str, Any]]) -> str:
        """统一的单元格内容恢复接口"""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.docx':
            return self.restore_cells_content_from_indexed_docx(file_path, restored_cells)
        elif file_ext in ['.xlsx', '.xls']:
            return self.restore_cells_content_from_indexed_excel(file_path, restored_cells)
        else:
            raise ValueError(f"Unsupported file format: {file_ext}")

