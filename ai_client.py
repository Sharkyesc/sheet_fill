import openai
import json
import re
from typing import List, Dict, Any
from colorama import Fore, Style
from config import Config
from rag_engine import RAGEngine

class AIClient:
    def __init__(self):
        self.client = openai.OpenAI(
            api_key=Config.OPENAI_API_KEY,
            base_url=Config.OPENAI_BASE_URL
        )
        self.model = Config.OPENAI_MODEL
        self.rag_engine = RAGEngine()

    def analyze_empty_fields_by_index(self, document_content: str) -> Dict[str, Any]:
        """Analyze numbered fields in the document and return info by index."""
        prompt = f"""
        请分析以下文档内容。对于每个标记了[index]的字段（例如[1], [2]等），
        请判断哪些字段需要填写，哪些字段不需要填写。
        
        判断标准：
        1. 如果字段已经有合适的内容，不需要填写，返回原内容
        2. 如果字段是空的或包含占位符（如"待填写"、"空白"等），需要填写
        3. 如果字段内容不完整或不合适，需要填写
        
        请返回JSON格式的结果：
        {{
          "fields": [
            {{
              "index": 1,
              "needs_filling": true/false,
              "description": "字段的含义描述，例如：性别",
              "suggested_content_type": "内容类型（例如：文本、日期、数字）",
              "original_content": "原始内容（如果不需要填写）"
            }}
          ]
        }}

        文档内容：
        {document_content}

        只返回有效的JSON格式。
        """
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一个智能文档理解助手，擅长分析表格中的标记字段并判断哪些需要填写。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1
            )
            print("AI raw response:", response)
            result = response.choices[0].message.content
            if result.strip().startswith("```"):
                match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", result, re.IGNORECASE)
                if match:
                    result = match.group(1)
            return json.loads(result)
        except Exception as e:
            print(f"按索引分析字段时出错: {e}")
            return {"fields": []}

    def search_with_rag(self, field_info: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        if self.rag_engine.get_index_stats()["document_count"] == 0:
            print("文本知识库为空，无法执行RAG搜索。")
            return []

        all_results = []
        for field in field_info:
            index = field.get("index")
            desc = field.get("description", "")
            print(f"\n{Fore.CYAN} 字段 [{index}] - {desc}{Style.RESET_ALL}")

            results = self.rag_engine.semantic_search([field], top_k=3)

            if results:
                for i, r in enumerate(results, 1):
                    content = r.get("content", "").strip()
                    score = r.get("similarity_score", 0)
                    print(f"{Fore.GREEN}  {i}. 相似度: {score:.3f}{Style.RESET_ALL}")
                    print(f"     内容: {content}")
            else:
                print(f"{Fore.RED}  未找到此字段的RAG结果{Style.RESET_ALL}")

            field["rag_evidence"] = results if results else []

            all_results.extend(results if results else [])

        return all_results

    def fill_document_with_rag(self, field_info: List[Dict[str, Any]], user_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        rag_context = self._build_rag_context(user_data)
        prompt = f"""
        根据以下字段信息（每个都有索引）和用户数据，为每个字段生成适当的内容。

        需要填充的字段:
        {json.dumps(field_info, ensure_ascii=False, indent=2)}

        用户数据（RAG检索结果）:
        {rag_context}

        请返回JSON格式:
        {{
            "field_answers": [
                {{
                    "index": 1,
                    "content": "要填充的内容",
                    "confidence": "high/medium/low",
                    "source": "数据来源"
                }}
            ]
        }}
        """
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一个使用RAG的专业文档填充助手。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2
            )
            print("AI原始响应:", response)
            result = response.choices[0].message.content
            if result.strip().startswith("```"):
                match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", result, re.IGNORECASE)
                if match:
                    result = match.group(1)
            return json.loads(result)
        except Exception as e:
            print(f"使用AI生成填充内容时出错: {e}")
            return {"field_answers": []}

    def _build_rag_context(self, user_data: List[Dict[str, Any]]) -> str:
        if not user_data:
            return "没有相关的用户数据。"
        return "\n".join(user.get("content", "") for user in user_data if user.get("content"))

    def get_rag_stats(self) -> Dict[str, Any]:
        return self.rag_engine.get_index_stats()

    def update_rag_index(self, new_documents: List[Dict[str, Any]]):
        self.rag_engine.update_index(new_documents)

    def fill_document(self, field_info: List[Dict[str, Any]], user_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        return self.fill_document_with_rag(field_info, user_data)
