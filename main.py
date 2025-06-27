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
        print(f"{Fore.GREEN}✓ Document filling system initialized{Style.RESET_ALL}")
        rag_stats = self.ai_client.get_rag_stats()
        print(f"{Fore.CYAN}RAG engine status: {rag_stats['status']}{Style.RESET_ALL}")

    def process_document(self, file_path: str) -> str:
        print(f"\n{Fore.CYAN}Start processing document: {file_path}{Style.RESET_ALL}")
        try:
            print(f"{Fore.YELLOW}Step 1: Find and number empty fields...{Style.RESET_ALL}")
            empty_fields, numbered_file = self.doc_processor.find_and_number_empty_fields_docx(file_path)
            if not empty_fields:
                print(f"{Fore.RED}No empty fields found, the document may already be filled{Style.RESET_ALL}")
                return file_path
            print(f"{Fore.GREEN}Found {len(empty_fields)} empty fields{Style.RESET_ALL}")

            highlighted_file = self.doc_processor.highlight_empty_fields_docx(numbered_file, empty_fields)
            print(f"{Fore.GREEN}✓ Highlighted saved as: {highlighted_file}{Style.RESET_ALL}")

            print(f"{Fore.YELLOW}Step 2: AI analyzes field descriptions...{Style.RESET_ALL}")
            doc_text = self.doc_processor.extract_document_content(highlighted_file)
            ai_response = self.ai_client.analyze_empty_fields_by_index(doc_text)
            if not ai_response.get("fields"):
                print(f"{Fore.RED}AI did not return any field descriptions.{Style.RESET_ALL}")
                return highlighted_file

            index_to_info = {f["index"]: f for f in ai_response["fields"]}
            for field in empty_fields:
                idx = field.get("index")
                info = index_to_info.get(idx)
                if info:
                    field["description"] = info.get("description", "")
                    field["suggested_content_type"] = info.get("suggested_content_type", "")
                else:
                    print(f"{Fore.YELLOW}No description matched for index {idx}{Style.RESET_ALL}")

            print(f"{Fore.YELLOW}Step 3: RAG search...{Style.RESET_ALL}")
            user_data = self.ai_client.search_with_rag(empty_fields)
            if not user_data:
                print(f"{Fore.YELLOW}No matching info found in knowledge base. Using all documents as fallback.{Style.RESET_ALL}")
                user_data = self.ai_client.rag_engine.documents
            print(f"{Fore.GREEN}✓ Found {len(user_data)} knowledge records{Style.RESET_ALL}")

            print(f"{Fore.YELLOW}Step 4: AI generates fill content...{Style.RESET_ALL}")
            fill_result = self.ai_client.fill_document_with_rag(empty_fields, user_data)
            answers = fill_result.get('field_answers', [])
            if not answers:
                print(f"{Fore.RED}AI could not generate fill content{Style.RESET_ALL}")
                return highlighted_file
            print(f"{Fore.GREEN}✓ AI generated {len(answers)} fill contents{Style.RESET_ALL}")

            indexed_fields = {f["index"]: f for f in empty_fields}
            for ans in answers:
                idx = ans["index"]
                field = indexed_fields.get(idx)
                if not field:
                    print(f"{Fore.YELLOW}Warning: index {idx} not found in empty_fields{Style.RESET_ALL}")
                    continue
                ans["table_index"] = field["table_index"]
                ans["row_index"] = field["row_index"]
                ans["col_index"] = field["col_index"]

            print(f"{Fore.YELLOW}Step 5: Fill document...{Style.RESET_ALL}")
            filled_file = self.doc_processor.fill_document(file_path, answers)
            print(f"{Fore.GREEN}✓ Document filled: {filled_file}{Style.RESET_ALL}")

            return filled_file
        except Exception as e:
            print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
            return file_path


def main():
    print(f"{Fore.CYAN}=== Intelligent Document Filler ==={Style.RESET_ALL}")
    filler = DocumentFiller()
    filler.ai_client.update_rag_index(filler.ai_client.rag_engine.load_txt_knowledge("./examples/sample_data.txt"))

    files = filler.doc_processor.list_docx_files()
    if not files:
        print(f"{Fore.YELLOW}No .docx files found in {Config.INPUT_DIR}{Style.RESET_ALL}")
        return
    for f in files:
        filler.process_document(f)

if __name__ == "__main__":
    main()