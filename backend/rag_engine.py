import os
import json
from typing import List, Dict, Any
import numpy as np
from sentence_transformers import SentenceTransformer
import faiss
from config import Config

class RAGEngine:
    def __init__(self, model_name: str = "all-MiniLM-L6-v2"):
        """Initialize RAG engine"""
        self.model_name = model_name
        self.model = SentenceTransformer(model_name)
        self.index = None
        self.documents = []
        self.document_embeddings = None

        self.index_path = os.path.normpath(os.path.join(Config.TEMP_DIR, "rag_index"))
        self.documents_path = os.path.normpath(os.path.join(Config.TEMP_DIR, "rag_documents.json"))
        os.makedirs(Config.TEMP_DIR, exist_ok=True)

    def l2_normalize(self, vectors):
        norms = np.linalg.norm(vectors, axis=1, keepdims=True)
        return vectors / (norms + 1e-10)

    def load_txt_knowledge(self, txt_path: str, chunk_size: int = 300, stride: int = 50) -> List[Dict[str, Any]]:
        """
        Load from plain text file and split into fixed-length chunks.
        """
        with open(txt_path, 'r', encoding='utf-8') as f:
            full_text = f.read()

        full_text = full_text.strip().replace('\r\n', '\n').replace('\r', '\n')
        paragraphs = full_text.split('\n\n')
        blocks = []
        block_id = 0
        for para in paragraphs:
            para = para.strip()
            if not para:
                continue
            for start in range(0, len(para), chunk_size - stride):
                chunk = para[start:start + chunk_size].strip()
                if chunk:
                    blocks.append({'id': block_id, 'content': chunk})
                    block_id += 1
        return blocks

    def add_documents(self, documents: List[Dict[str, Any]]):
        """Add documents and build index"""
        if not documents:
            print("No documents to add")
            return
            
        print(f"Adding {len(documents)} documents to RAG index...")

        text_chunks = [doc.get('content', '') for doc in documents]
        embeddings = self.model.encode(text_chunks, show_progress_bar=True)
        embeddings = self.l2_normalize(embeddings)

        dimension = embeddings.shape[1]
        self.index = faiss.IndexFlatIP(dimension)
        self.index.add(embeddings.astype('float32'))

        self.documents = documents
        self.document_embeddings = embeddings
        self._save_index()
        print(f"✓ RAG index created with {len(documents)} documents")

    def search(self, query: str, top_k: int = 5) -> List[Dict[str, Any]]:
        """Semantic search for relevant documents"""
        if self.index is None:
            print("RAG index not initialized, attempting to load...")
            if not self._load_index():
                return []

        query_embedding = self.model.encode([query])
        query_embedding = self.l2_normalize(query_embedding)
        scores, indices = self.index.search(query_embedding.astype('float32'), top_k)

        results = []
        for i, (score, idx) in enumerate(zip(scores[0], indices[0])):
            if idx < len(self.documents):
                result = self.documents[idx].copy()
                result['similarity_score'] = float(score)
                result['rank'] = i + 1
                results.append(result)

        return results

    def semantic_search(self, field_info: List[Dict[str, Any]], top_k: int = 3) -> List[Dict[str, Any]]:
        """Semantic search based on field information"""
        if not field_info:
            return []

        queries = []
        for field in field_info:
            field_type = field.get('field_type', '')
            description = field.get('description', '')
            suggested_type = field.get('suggested_content_type', '')

            if '姓名' in field_type or 'name' in field_type.lower():
                queries.append("姓名 名字 员工姓名")
            elif '邮箱' in field_type or 'email' in field_type.lower():
                queries.append("邮箱 email 电子邮件")
            elif '电话' in field_type or 'phone' in field_type.lower():
                queries.append("电话 手机 联系方式")
            elif '地址' in field_type or 'address' in field_type.lower():
                queries.append("地址 住址 工作地址")
            elif '公司' in field_type or 'company' in field_type.lower():
                queries.append("公司 企业 工作单位")
            elif '职位' in field_type or 'position' in field_type.lower():
                queries.append("职位 岗位 职务")
            elif '部门' in field_type or 'department' in field_type.lower():
                queries.append("部门 科室 团队")
            elif '技能' in field_type or 'skill' in field_type.lower():
                queries.append("技能 技术 能力")
            elif '教育' in field_type or 'education' in field_type.lower():
                queries.append("教育 学历 学校")
            elif '经验' in field_type or 'experience' in field_type.lower():
                queries.append("经验 工作经历 履历")
            else:
                queries.append(" ".join([description, suggested_type]).strip())

        all_results = []
        for query in queries:
            if query.strip():
                results = self.search(query, top_k)
                all_results.extend(results)

        unique_results = self._deduplicate_results(all_results)
        return unique_results[:top_k]

    def _deduplicate_results(self, results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Deduplicate and sort by similarity"""
        seen_ids = set()
        unique_results = []

        for result in results:
            result_id = result.get('id')
            if result_id not in seen_ids:
                seen_ids.add(result_id)
                unique_results.append(result)

        unique_results.sort(key=lambda x: x.get('similarity_score', 0), reverse=True)
        return unique_results

    def _save_index(self):
        """Save index and documents"""
        try:
            os.makedirs(os.path.dirname(self.index_path), exist_ok=True)
            print(f"Saving FAISS index to: {self.index_path}")
            faiss.write_index(self.index, self.index_path)

            print(f"Saving documents to: {self.documents_path}")
            with open(self.documents_path, 'w', encoding='utf-8') as f:
                json.dump(self.documents, f, ensure_ascii=False, indent=2)

            print(f"✓ RAG index successfully saved to {Config.TEMP_DIR}")
        except Exception as e:
            print(f"Failed to save RAG index: {e}")
            print(f"Index path: {self.index_path}")
            print(f"Documents path: {self.documents_path}")

    def _load_index(self) -> bool:
        """Load existing index"""
        try:
            if not os.path.exists(self.index_path) or not os.path.exists(self.documents_path):
                return False

            self.index = faiss.read_index(self.index_path)
            with open(self.documents_path, 'r', encoding='utf-8') as f:
                self.documents = json.load(f)

            print(f"RAG index loaded successfully, contains {len(self.documents)} documents")
            return True
        except Exception as e:
            print(f"Failed to load RAG index: {e}")
            return False

    def update_index(self, new_documents: List[Dict[str, Any]]):
        """Merge and rebuild index"""
        if self.index is None:
            self.add_documents(new_documents)
        else:
            all_documents = self.documents + new_documents
            self.add_documents(all_documents)

    def get_index_stats(self) -> Dict[str, Any]:
        """Get index statistics"""
        if self.index is None:
            return {"status": "Not initialized", "document_count": 0}
        return {
            "status": "Initialized",
            "document_count": len(self.documents),
            "index_size": self.index.ntotal,
            "model_name": self.model_name
        }
