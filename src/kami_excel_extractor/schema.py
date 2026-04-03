from pydantic import BaseModel, Field
from typing import List, Dict, Any, Optional, Union

class ExtractionResult(BaseModel):
    """LLMからの抽出結果を個別に保持するモデル"""
    data: Optional[Union[List[Any], Dict[str, Any]]] = None
    _raw_json: Optional[str] = None
    error: Optional[str] = None

class SheetData(BaseModel):
    """シートごとの解析結果スキーマ（柔軟な構造を許容）"""
    metadata: Optional[Dict[str, Any]] = Field(default_factory=dict)
    sections: Optional[List[Dict[str, Any]]] = Field(default_factory=list)
    data: Optional[Union[List[Any], Dict[str, Any]]] = Field(default_factory=list)
    errors: Optional[List[str]] = Field(default_factory=list)

class FullExtraction(BaseModel):
    """Excel全体の抽出データ"""
    sheets: Dict[str, SheetData]
