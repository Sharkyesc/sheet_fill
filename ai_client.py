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
        Please analyze the following document content. For each field labeled with [index] (e.g., [1], [2], etc.),
        If you think a field can be empty, just leave it empty.
        return a JSON list with the following structure:
        
        {{
          "fields": [
            {{
              "index": 1,
              "description": "The meaning of this field, e.g., Gender",
              "suggested_content_type": "Type of content (e.g., text, date, number)"
            }}
          ]
        }}

        Document content:
        {document_content}

        Only return valid JSON.
        """
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are a smart document understanding assistant, good at understanding labeled fields in tables."},
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
            print(f"Error analyzing fields by index: {e}")
            return {"fields": []}

    def search_with_rag(self, field_info: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        if self.rag_engine.get_index_stats()["document_count"] == 0:
            print("Txt knowledge base is empty, cannot perform RAG search.")
            return []

        all_results = []
        for field in field_info:
            index = field.get("index")
            desc = field.get("description", "")
            print(f"\n{Fore.CYAN} Field [{index}] - {desc}{Style.RESET_ALL}")

            results = self.rag_engine.semantic_search([field], top_k=3)

            if results:
                for i, r in enumerate(results, 1):
                    content = r.get("content", "").strip()
                    score = r.get("similarity_score", 0)
                    print(f"{Fore.GREEN}  {i}. Score: {score:.3f}{Style.RESET_ALL}")
                    print(f"     Content: {content}")
            else:
                print(f"{Fore.RED}  No RAG result found for this field{Style.RESET_ALL}")

            field["rag_evidence"] = results if results else []

            all_results.extend(results if results else [])

        return all_results

    def fill_document_with_rag(self, field_info: List[Dict[str, Any]], user_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        rag_context = self._build_rag_context(user_data)
        prompt = f"""
        Based on the following field information (each with an index) and user data, generate appropriate content for each field.

        Fields to be filled:
        {json.dumps(field_info, ensure_ascii=False, indent=2)}

        User data (RAG retrieval results):
        {rag_context}

        Please return in JSON:
        {{
            "field_answers": [
                {{
                    "index": 1,
                    "content": "content to fill",
                    "confidence": "high/medium/low",
                    "source": "data source"
                }}
            ]
        }}
        """
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are a professional document filling assistant using RAG."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2
            )
            print("AI raw response:", response)
            result = response.choices[0].message.content
            if result.strip().startswith("```"):
                match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", result, re.IGNORECASE)
                if match:
                    result = match.group(1)
            return json.loads(result)
        except Exception as e:
            print(f"Error generating fill content with AI: {e}")
            return {"field_answers": []}

    def _build_rag_context(self, user_data: List[Dict[str, Any]]) -> str:
        if not user_data:
            return "No relevant user data."
        return "\n".join(user.get("content", "") for user in user_data if user.get("content"))

    def get_rag_stats(self) -> Dict[str, Any]:
        return self.rag_engine.get_index_stats()

    def update_rag_index(self, new_documents: List[Dict[str, Any]]):
        self.rag_engine.update_index(new_documents)

    def fill_document(self, field_info: List[Dict[str, Any]], user_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        return self.fill_document_with_rag(field_info, user_data)
