import os
from typing import List, Dict, Any
from colorama import init, Fore, Style

from config import Config
from ai_client import AIClient
from document_processor import DocumentProcessor

# Initialize colorama
init()

class DocumentFiller:
    def __init__(self):
        Config.create_directories()
        self.ai_client = AIClient()
        self.doc_processor = DocumentProcessor()
        print(f"{Fore.GREEN}✓ 文档填充系统已初始化{Style.RESET_ALL}")
        rag_stats = self.ai_client.get_rag_stats()
        print(f"{Fore.CYAN}RAG引擎状态: {rag_stats['status']}{Style.RESET_ALL}")

    def process_document(self, file_path: str) -> str:
        print(f"\n{Fore.CYAN}开始处理文档: {file_path}{Style.RESET_ALL}")
        try:
            print(f"{Fore.YELLOW}步骤1: 为所有字段添加索引...{Style.RESET_ALL}")
            all_fields, numbered_file = self.doc_processor.find_and_number_empty_fields_docx(file_path)
            if not all_fields:
                print(f"{Fore.RED}未找到任何字段{Style.RESET_ALL}")
                return file_path
            print(f"{Fore.GREEN}找到 {len(all_fields)} 个字段{Style.RESET_ALL}")

            print(f"{Fore.YELLOW}步骤2: AI分析字段并判断哪些需要填写...{Style.RESET_ALL}")
            doc_text = self.doc_processor.extract_document_content(numbered_file)
            ai_response = self.ai_client.analyze_empty_fields_by_index(doc_text)
            if not ai_response.get("fields"):
                print(f"{Fore.RED}AI未返回任何字段描述{Style.RESET_ALL}")
                return numbered_file

            # 根据AI判断结果过滤字段
            index_to_ai_info = {f["index"]: f for f in ai_response["fields"]}
            fields_to_fill = []
            fields_to_restore = []

            for field in all_fields:
                idx = field.get("index")
                ai_info = index_to_ai_info.get(idx)
                if ai_info:
                    field["description"] = ai_info.get("description", "")
                    field["suggested_content_type"] = ai_info.get("suggested_content_type", "")
                    
                    if ai_info.get("needs_filling", True):
                        fields_to_fill.append(field)
                        print(f"{Fore.CYAN}字段 [{idx}] 需要填写: {ai_info.get('description', '')}{Style.RESET_ALL}")
                    else:
                        fields_to_restore.append({
                            "index": idx,
                            "table_index": field["table_index"],
                            "row_index": field["row_index"],
                            "col_index": field["col_index"],
                            "original_content": ai_info.get("original_content", field["text"])
                        })
                        print(f"{Fore.GREEN}字段 [{idx}] 不需要填写，将恢复原内容: {ai_info.get('original_content', field['text'])}{Style.RESET_ALL}")
                else:
                    print(f"{Fore.YELLOW}未找到索引 {idx} 的AI描述，默认需要填写{Style.RESET_ALL}")
                    fields_to_fill.append(field)

            if not fields_to_fill:
                print(f"{Fore.GREEN}所有字段都不需要填写，直接恢复原内容{Style.RESET_ALL}")
                # 恢复所有字段的原内容
                self.doc_processor.restore_original_content(numbered_file, fields_to_restore)
                return numbered_file

            print(f"{Fore.GREEN}需要填写的字段: {len(fields_to_fill)} 个{Style.RESET_ALL}")
            print(f"{Fore.GREEN}不需要填写的字段: {len(fields_to_restore)} 个{Style.RESET_ALL}")

            # 高亮需要填写的字段
            highlighted_file = self.doc_processor.highlight_empty_fields_docx(numbered_file, fields_to_fill)
            print(f"{Fore.GREEN}✓ 高亮文件保存为: {highlighted_file}{Style.RESET_ALL}")

            print(f"{Fore.YELLOW}步骤3: RAG搜索...{Style.RESET_ALL}")
            user_data = self.ai_client.search_with_rag(fields_to_fill)
            if not user_data:
                print(f"{Fore.YELLOW}在知识库中未找到匹配信息。使用所有文档作为备用。{Style.RESET_ALL}")
                user_data = self.ai_client.rag_engine.documents
            print(f"{Fore.GREEN}✓ 找到 {len(user_data)} 条知识记录{Style.RESET_ALL}")

            print(f"{Fore.YELLOW}步骤4: AI生成填充内容...{Style.RESET_ALL}")
            fill_result = self.ai_client.fill_document_with_rag(fields_to_fill, user_data)
            answers = fill_result.get('field_answers', [])
            if not answers:
                print(f"{Fore.RED}AI无法生成填充内容{Style.RESET_ALL}")
                return highlighted_file
            print(f"{Fore.GREEN}✓ AI生成了 {len(answers)} 个填充内容{Style.RESET_ALL}")

            # 合并需要填写的答案和需要恢复的内容
            all_answers = answers + fields_to_restore

            indexed_fields = {f["index"]: f for f in all_fields}
            for ans in all_answers:
                idx = ans["index"]
                field = indexed_fields.get(idx)
                if not field:
                    print(f"{Fore.YELLOW}警告: 索引 {idx} 在字段中未找到{Style.RESET_ALL}")
                    continue
                ans["table_index"] = field["table_index"]
                ans["row_index"] = field["row_index"]
                ans["col_index"] = field["col_index"]

            print(f"{Fore.YELLOW}步骤5: 填充文档...{Style.RESET_ALL}")
            filled_file = self.doc_processor.fill_document(file_path, all_answers)
            print(f"{Fore.GREEN}✓ 文档已填充: {filled_file}{Style.RESET_ALL}")

            return filled_file
        except Exception as e:
            print(f"{Fore.RED}错误: {e}{Style.RESET_ALL}")
            return file_path


def main():
    print(f"{Fore.CYAN}=== 智能文档填充系统 ==={Style.RESET_ALL}")
    filler = DocumentFiller()
    filler.ai_client.update_rag_index(filler.ai_client.rag_engine.load_txt_knowledge("./examples/sample_data.txt"))

    files = filler.doc_processor.list_docx_files()
    if not files:
        print(f"{Fore.YELLOW}在 {Config.INPUT_DIR} 中未找到 .docx 文件{Style.RESET_ALL}")
        return
    for f in files:
        filler.process_document(f)

if __name__ == "__main__":
    main()