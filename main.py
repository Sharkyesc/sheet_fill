import os
import sys
import time
import argparse
from typing import List, Dict, Any
from colorama import init, Fore, Style

from config import Config
from ai_client import AIClient
from document_processor import DocumentProcessor
from pdf_processor import PDFProcessor
from monitor import SystemMonitor

# Initialize colorama
init()

class Logger:
    def __init__(self, log_file_path: str):
        self.log_file_path = log_file_path
        self.terminal = sys.stdout
        
    def write(self, message):
        self.terminal.write(message)
        self.terminal.flush()
        
        clean_message = self.remove_color_codes(message)
        with open(self.log_file_path, 'a', encoding='utf-8') as f:
            f.write(clean_message)
            f.flush()
    
    def flush(self):
        self.terminal.flush()
    
    def remove_color_codes(self, text):
        import re
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        return ansi_escape.sub('', text)

timestamp = int(time.time())
log_filename = f"document_fill_log_{timestamp}.txt"
log_file_path = os.path.join(Config.TEMP_DIR, log_filename)
logger = Logger(log_file_path)
sys.stdout = logger

class DocumentFiller:
    def __init__(self, enable_monitoring: bool = True, monitor_interval: int = 100):
        Config.create_directories()
        self.ai_client = AIClient()
        self.doc_processor = DocumentProcessor()
        self.pdf_processor = PDFProcessor()
        
        self.monitor = None
        self.enable_monitoring = enable_monitoring
        self.monitor_interval = monitor_interval
        if enable_monitoring:
            self.monitor = SystemMonitor(interval = monitor_interval)
            print(f"{Fore.GREEN}✓ Document filling system initialized with monitoring capability{Style.RESET_ALL}")
        else:
            print(f"{Fore.GREEN}✓ Document filling system initialized{Style.RESET_ALL}")
            
        print(f"{Fore.CYAN}Log file: {log_file_path}{Style.RESET_ALL}")
        rag_stats = self.ai_client.get_rag_stats()
        print(f"{Fore.CYAN}RAG engine status: {rag_stats['status']}{Style.RESET_ALL}")

    def process_document(self, file_path: str) -> str:
        print(f"\n{Fore.CYAN}Start processing document: {file_path}{Style.RESET_ALL}")
        
        if self.enable_monitoring and self.monitor:
            self.monitor.start_monitoring()
        try:
            print(f"{Fore.YELLOW}Step 1: Number all fields in the document...{Style.RESET_ALL}")
            # 使用统一的字段标记接口
            all_fields, numbered_file = self.doc_processor.find_and_number_all_fields(file_path)
            if not all_fields:
                print(f"{Fore.RED}No fields found in the document.{Style.RESET_ALL}")
                return file_path
            print(f"{Fore.GREEN}Found {len(all_fields)} fields in the document{Style.RESET_ALL}")

            print(f"{Fore.YELLOW}Step 2: Convert document to PDF and generate page screenshots...{Style.RESET_ALL}")
            try:
                pdf_result = self.pdf_processor.process_document_with_images(numbered_file)
                pdf_result = self.pdf_processor.process_document_with_images(numbered_file)
                page_images = pdf_result['page_images']
                pdf_path = pdf_result['pdf_path']
                print(f"{Fore.GREEN}✓ PDF conversion and screenshots completed, total {len(page_images)} pages{Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.RED}PDF processing failed: {e}, will use text-only analysis{Style.RESET_ALL}")
                page_images = []
                pdf_path = None

            print(f"{Fore.YELLOW}Step 3: AI analyzes fields (combining images and text content)...{Style.RESET_ALL}")
            doc_text = self.doc_processor.extract_document_content(numbered_file)
            
            if page_images:
                ai_response = self.ai_client.analyze_empty_fields_with_images(doc_text, page_images)
                print(f"{Fore.GREEN}✓ Analysis with images and text content completed{Style.RESET_ALL}")
            else:
                ai_response = self.ai_client.analyze_empty_fields_by_index(doc_text)
                print(f"{Fore.GREEN}✓ Text-only content analysis completed{Style.RESET_ALL}")
            
            if not ai_response.get("fields_to_fill"):
                print(f"{Fore.RED}AI did not return any field descriptions.{Style.RESET_ALL}")
                return numbered_file
            described_fields = ai_response["fields_to_fill"]

            print(f"\n{Fore.CYAN}=== AI Field Analysis Results ==={Style.RESET_ALL}")
            for field in described_fields:
                print(f"Field [{field.get('index', 'N/A')}] | Description: {field.get('description', 'N/A')} | Content Type: {field.get('suggested_content_type', 'N/A')}")
            if ai_response.get('restored_cells'):
                print(f"\n{Fore.BLUE}=== AI Restored Cells ==={Style.RESET_ALL}")
                for cell in ai_response['restored_cells']:
                    print(f"Restored Field [{cell.get('index', 'N/A')}] | Content: {cell.get('restored_content', 'N/A')}")

            print(f"{Fore.YELLOW}Step 4: RAG search for each field...{Style.RESET_ALL}")
            for field in described_fields:
                print(f"\n{Fore.CYAN}--- Field [{field['index']}] ---{Style.RESET_ALL}")
                print(f"{Fore.WHITE}Description: {field.get('description', 'N/A')}{Style.RESET_ALL}")
                print(f"{Fore.WHITE}Content Type: {field.get('suggested_content_type', 'N/A')}{Style.RESET_ALL}")
                
                rag_results = self.ai_client.rag_engine.semantic_search([field], top_k=3)
                field["rag_evidence"] = rag_results
                
                if rag_results:
                    print(f"{Fore.GREEN}✓ Found {len(rag_results)} RAG matches:{Style.RESET_ALL}")
                    for i, result in enumerate(rag_results, 1):
                        score = result.get('similarity_score', 0)
                        content = result.get('content', 'N/A')
                        print(f"  {i}. Score: {score:.3f}")
                        print(f"     Content: {content[:100]}{'...' if len(content) > 100 else ''}")
                else:
                    print(f"{Fore.RED}✗ No RAG matches found{Style.RESET_ALL}")
            
            print(f"\n{Fore.GREEN}✓ RAG evidence added to all described fields{Style.RESET_ALL}")

            print(f"{Fore.YELLOW}Step 5: AI makes final fill/restore decision...{Style.RESET_ALL}")
            final_decision = self.ai_client.final_fill_decision(all_fields, described_fields)
            filled_cells = final_decision.get("filled_cells", [])
            restored_cells = final_decision.get("restored_cells", [])
            
            print(f"\n{Fore.CYAN}=== AI Decision Results ==={Style.RESET_ALL}")
            if filled_cells:
                print(f"\n{Fore.GREEN}✓ Cells to be filled ({len(filled_cells)}):{Style.RESET_ALL}")
                for cell in filled_cells:
                    idx = cell.get('index')
                    content = cell.get('content', 'N/A')
                    print(f"  Field [{idx}]: {content}")
            else:
                print(f"\n{Fore.YELLOW}No cells to be filled{Style.RESET_ALL}")
                
            if restored_cells:
                print(f"\n{Fore.BLUE}✓ Cells to be restored ({len(restored_cells)}):{Style.RESET_ALL}")
                for cell in restored_cells:
                    idx = cell.get('index')
                    content = cell.get('restored_content', 'N/A')
                    print(f"  Field [{idx}]: {content}")
            else:
                print(f"\n{Fore.YELLOW}No cells to be restored{Style.RESET_ALL}")
            
            if not filled_cells and not restored_cells:
                print(f"{Fore.YELLOW}No cells to fill or restore according to AI.{Style.RESET_ALL}")
                return numbered_file

            restored_file = numbered_file
            if restored_cells:
                print(f"{Fore.YELLOW}Step 6: Restoring {len(restored_cells)} cells...{Style.RESET_ALL}")
                indexed_fields = {f["index"]: f for f in all_fields}
                for cell_info in restored_cells:
                    idx = cell_info.get("index")
                    field = indexed_fields.get(idx)
                    if field:
                        cell_info["original_format"] = field.get("original_format")
                # 使用统一的恢复接口
                restored_file = self.doc_processor.restore_cells_content_from_indexed(numbered_file, restored_cells)
                print(f"{Fore.GREEN}✓ Cells restored: {restored_file}{Style.RESET_ALL}")

            if filled_cells:
                print(f"{Fore.YELLOW}Filling cells with AI-generated content...{Style.RESET_ALL}")
                # 根据文件类型准备填充数据
                file_ext = os.path.splitext(file_path)[1].lower()
                indexed_fields = {f["index"]: f for f in all_fields}
                
                for ans in filled_cells:
                    idx = ans["index"]
                    field = indexed_fields.get(idx)
                    if not field:
                        print(f"{Fore.YELLOW}Warning: index {idx} not found in all_fields{Style.RESET_ALL}")
                        continue
                    
                    # 为不同文件类型设置不同的字段信息
                    if file_ext == '.docx':
                        # Word文档字段映射
                        ans["table_index"] = field["table_index"]
                        ans["row_index"] = field["row_index"]
                        ans["col_index"] = field["col_index"]
                    elif file_ext in ['.xlsx', '.xls']:
                        # Excel文档字段映射
                        ans["sheet_name"] = field["sheet_name"]
                        ans["row_index"] = field["row_index"]
                        ans["col_index"] = field["col_index"]
                    
                    ans["original_format"] = field.get("original_format")
                
                filled_file = self.doc_processor.fill_document(restored_file, filled_cells)
                print(f"{Fore.GREEN}✓ Document filled: {filled_file}{Style.RESET_ALL}")
                
                if pdf_path:
                    self.pdf_processor.cleanup_temp_files(pdf_path)
                
                return filled_file
            else:
                print(f"{Fore.BLUE}No cells need to be filled, returning restored file: {restored_file}{Style.RESET_ALL}")
                if pdf_path:
                    self.pdf_processor.cleanup_temp_files(pdf_path)
                return restored_file
                
        except Exception as e:
            print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
            return file_path
    
    def stop_monitoring_and_generate_report(self):
        if self.monitor and self.enable_monitoring:
            print(f"\n{Fore.CYAN}=== Stop monitoring and generate report ==={Style.RESET_ALL}")
            self.monitor.stop_monitoring()
            
            self.monitor.print_summary()
            data_file = self.monitor.save_data()
            chart_file = self.monitor.generate_charts()
            
            print(f"{Fore.GREEN}✓ Monitoring data saved to: {data_file}{Style.RESET_ALL}")
            print(f"{Fore.GREEN}✓ Monitoring chart saved to: {chart_file}{Style.RESET_ALL}")
            
            return data_file, chart_file
        else:
            print(f"{Fore.YELLOW}Monitoring not enabled{Style.RESET_ALL}")
            return None, None

def main():
    parser = argparse.ArgumentParser(description="Intelligent Sheet Filling System")
    parser.add_argument('--knowledge', type=str, help='Knowledge file path (single text file)')
    parser.add_argument('--knowledge-files', type=str, nargs='+', help='Knowledge files path (support multiple files: txt/doc/docx/pdf)')
    parser.add_argument('--forms', type=str, nargs='+', required=True, help='Form file path (support multiple files: .docx/.xlsx/.xls/.doc)')
    parser.add_argument('--no-monitor', action='store_true', help='Disable system monitoring')
    parser.add_argument('--monitor-interval', type=int, default=100, help='Monitoring interval in ms (default: 100)')
    args = parser.parse_args()

    if not args.knowledge and not args.knowledge_files:
        print(f"{Fore.RED}Must provide --knowledge or --knowledge-files parameter{Style.RESET_ALL}")
        return
    
    if args.knowledge and args.knowledge_files:
        print(f"{Fore.RED}Cannot use both --knowledge and --knowledge-files parameters{Style.RESET_ALL}")
        return

    for f in args.forms:
        if not os.path.exists(f):
            print(f"{Fore.RED}Form file not found: {f}{Style.RESET_ALL}")
            return

    print(f"{Fore.CYAN}=== Intelligent Document Filler (Word & Excel Support) ==={Style.RESET_ALL}")
    enable_monitoring = not args.no_monitor
    filler = DocumentFiller(enable_monitoring=enable_monitoring, monitor_interval=args.monitor_interval)

    ai_client = AIClient()
    
    if args.knowledge_files:
        print(f"{Fore.CYAN}Building RAG knowledge base from files...{Style.RESET_ALL}")
        
        for f in args.knowledge_files:
            if not os.path.exists(f):
                print(f"{Fore.RED}Knowledge file not found: {f}{Style.RESET_ALL}")
                return
        
        success = ai_client.build_rag_from_files(args.knowledge_files)
        if not success:
            print(f"{Fore.RED}Failed to build RAG knowledge base from files{Style.RESET_ALL}")
            return
            
    else:
        if not os.path.exists(args.knowledge):
            print(f"{Fore.RED}Knowledge file not found: {args.knowledge}{Style.RESET_ALL}")
            return
            
        print(f"{Fore.CYAN}Building RAG knowledge base from text file...{Style.RESET_ALL}")
        with open(args.knowledge, 'r', encoding='utf-8') as f:
            full_text = f.read()
        llm_chunks = ai_client.split_text_with_llm(full_text, max_chunk_length=300)
        if llm_chunks and len(llm_chunks) > 0:
            documents = [ {'id': i, 'content': chunk} for i, chunk in enumerate(llm_chunks) ]
            filler.ai_client.update_rag_index(documents)
        else:
            documents = filler.ai_client.rag_engine.load_txt_knowledge(args.knowledge)
            filler.ai_client.update_rag_index(documents)

    files = args.forms
    for i, f in enumerate(files):
        ext = os.path.splitext(f)[1].lower()
        if ext == '.doc':
            # 将.doc转换为.docx
            docx_path = f + 'x' if not f.endswith('.docx') else f
            print(f"{Fore.YELLOW}Converting {f} to {docx_path}{Style.RESET_ALL}")
            f = filler.doc_processor.convert_doc_to_docx(f, docx_path)
            files[i] = f
        elif ext in ['.xls', '.xlsx', '.docx']:
            # 支持的格式，直接处理
            print(f"{Fore.CYAN}Processing {ext.upper()} file: {os.path.basename(f)}{Style.RESET_ALL}")
        else:
            print(f"{Fore.YELLOW}Warning: Unsupported file format {ext} for file: {f}{Style.RESET_ALL}")
            continue
        
        filler.process_document(f)
    
    print(f"\n{Fore.GREEN}=== Processing Complete ==={Style.RESET_ALL}")
    print(f"{Fore.CYAN}All output has been saved to: {log_file_path}{Style.RESET_ALL}")
    
    data_file, chart_file = filler.stop_monitoring_and_generate_report()
    
    sys.stdout = logger.terminal

if __name__ == "__main__":
    main()