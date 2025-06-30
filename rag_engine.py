import os
import json
import pickle
from typing import List, Dict, Any, Tuple
import numpy as np
from sentence_transformers import SentenceTransformer
import faiss
from sklearn.metrics.pairwise import cosine_similarity
from config import Config

class RAGEngine:
    def __init__(self, model_name: str = "all-MiniLM-L6-v2"):
        """初始化RAG引擎"""
        self.model_name = model_name
        self.model = SentenceTransformer(model_name)
        self.index = None
        self.documents = []
        self.document_embeddings = None
        
        # 确保路径格式正确
        self.index_path = os.path.normpath(os.path.join(Config.TEMP_DIR, "rag_index.pkl"))
        self.documents_path = os.path.normpath(os.path.join(Config.TEMP_DIR, "rag_documents.json"))
        
        os.makedirs(Config.TEMP_DIR, exist_ok=True)
    
    def load_txt_knowledge(self, txt_path: str):
        docs = []
        with open(txt_path, 'r', encoding='utf-8') as f:
            for i, line in enumerate(f):
                line = line.strip()
                if line:
                    docs.append({'id': i, 'content': line})
        return docs
    
    def add_documents(self, documents: List[Dict[str, Any]]):
        """添加文档到RAG索引"""
        print(f"Adding {len(documents)} documents to RAG index...")
        # 支持txt知识库：如果只有content字段，直接用content
        text_chunks = []
        for doc in documents:
            if 'content' in doc:
                chunk = doc['content']
            else:
                chunk = self._document_to_text_chunk(doc)
            text_chunks.append(chunk)
        embeddings = self.model.encode(text_chunks, show_progress_bar=True)
        dimension = embeddings.shape[1]
        self.index = faiss.IndexFlatIP(dimension)
        self.index.add(embeddings.astype('float32'))
        self.documents = documents
        self.document_embeddings = embeddings
        self._save_index()
        print(f"✓ RAG index created with {len(documents)} documents")
    
    def _document_to_text_chunk(self, doc: Dict[str, Any]) -> str:
        """将文档转换为文本块"""
        chunks = []
        
        # 基本信息
        if doc.get('name'):
            chunks.append(f"姓名: {doc['name']}")
        if doc.get('email'):
            chunks.append(f"邮箱: {doc['email']}")
        if doc.get('phone'):
            chunks.append(f"电话: {doc['phone']}")
        if doc.get('address'):
            chunks.append(f"地址: {doc['address']}")
        
        # 工作信息
        if doc.get('company'):
            chunks.append(f"公司: {doc['company']}")
        if doc.get('position'):
            chunks.append(f"职位: {doc['position']}")
        if doc.get('department'):
            chunks.append(f"部门: {doc['department']}")
        if doc.get('employee_id'):
            chunks.append(f"员工ID: {doc['employee_id']}")
        if doc.get('hire_date'):
            chunks.append(f"入职日期: {doc['hire_date']}")
        if doc.get('salary'):
            chunks.append(f"薪资: {doc['salary']}")
        
        # 技能信息
        if doc.get('skills'):
            if isinstance(doc['skills'], list):
                skills_text = ', '.join(doc['skills'])
            else:
                skills_text = str(doc['skills'])
            chunks.append(f"技能: {skills_text}")
        
        # 教育背景
        if doc.get('education'):
            if isinstance(doc['education'], list):
                for edu in doc['education']:
                    if isinstance(edu, dict):
                        edu_text = f"{edu.get('degree', '')} - {edu.get('school', '')} - {edu.get('major', '')}"
                        chunks.append(f"教育: {edu_text}")
            else:
                chunks.append(f"教育: {doc['education']}")
        
        # 工作经验
        if doc.get('experience'):
            if isinstance(doc['experience'], list):
                for exp in doc['experience']:
                    if isinstance(exp, dict):
                        exp_text = f"{exp.get('company', '')} - {exp.get('position', '')} - {exp.get('duration', '')}"
                        chunks.append(f"经验: {exp_text}")
            else:
                chunks.append(f"经验: {doc['experience']}")
        
        return " | ".join(chunks)
    
    def search(self, query: str, top_k: int = 5) -> List[Dict[str, Any]]:
        """搜索相关文档"""
        if self.index is None:
            print("RAG索引未初始化，尝试加载...")
            if not self._load_index():
                return []
        
        query_embedding = self.model.encode([query])
        
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
        """基于字段信息的语义搜索"""
        if not field_info:
            return []
        
        # 构建搜索查询
        queries = []
        for field in field_info:
            field_type = field.get('field_type', '')
            description = field.get('description', '')
            suggested_type = field.get('suggested_content_type', '')
            
            # 根据字段类型构建查询
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
                # 使用描述和上下文构建查询
                query_parts = [description, suggested_type]
                queries.append(" ".join([part for part in query_parts if part]))
        
        # 执行搜索
        all_results = []
        for query in queries:
            if query.strip():
                results = self.search(query, top_k)
                all_results.extend(results)
        
        # 去重并排序
        unique_results = self._deduplicate_results(all_results)
        return unique_results[:top_k]
    
    def _deduplicate_results(self, results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """去重结果"""
        seen_ids = set()
        unique_results = []
        
        for result in results:
            result_id = result.get('id')
            if result_id not in seen_ids:
                seen_ids.add(result_id)
                unique_results.append(result)
        
        # 按相似度排序
        unique_results.sort(key=lambda x: x.get('similarity_score', 0), reverse=True)
        return unique_results
    
    def _save_index(self):
        """保存索引和文档"""
        try:
            os.makedirs(os.path.dirname(self.index_path), exist_ok=True)
            
            print(f"正在保存FAISS索引到: {self.index_path}")
            faiss.write_index(self.index, self.index_path)
            
            print(f"正在保存文档到: {self.documents_path}")
            with open(self.documents_path, 'w', encoding='utf-8') as f:
                json.dump(self.documents, f, ensure_ascii=False, indent=2)
            
            print(f"✓ RAG索引已成功保存到 {Config.TEMP_DIR}")
        except Exception as e:
            print(f"保存RAG索引失败: {e}")
            print(f"索引路径: {self.index_path}")
            print(f"文档路径: {self.documents_path}")
    
    def _load_index(self) -> bool:
        """加载索引和文档"""
        try:
            if not os.path.exists(self.index_path) or not os.path.exists(self.documents_path):
                return False
            
            # 加载FAISS索引
            self.index = faiss.read_index(self.index_path)
            
            # 加载文档
            with open(self.documents_path, 'r', encoding='utf-8') as f:
                self.documents = json.load(f)
            
            print(f"RAG索引加载完成，包含 {len(self.documents)} 个文档")
            return True
        except Exception as e:
            print(f"加载RAG索引失败: {e}")
            return False
    
    def update_index(self, new_documents: List[Dict[str, Any]]):
        """更新索引"""
        if self.index is None:
            self.add_documents(new_documents)
        else:
            # 合并文档
            all_documents = self.documents + new_documents
            self.add_documents(all_documents)
    
    def get_index_stats(self) -> Dict[str, Any]:
        """获取索引统计信息"""
        if self.index is None:
            return {"status": "未初始化", "document_count": 0}
        
        return {
            "status": "已初始化",
            "document_count": len(self.documents),
            "index_size": self.index.ntotal,
            "model_name": self.model_name
        } 