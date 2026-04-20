from pydantic import BaseModel, Field, ConfigDict
from typing import List, Dict, Any, Optional, Union

class ExtractionOptions(BaseModel):
    """データ抽出時のオプション"""
    model: Optional[str] = None
    system_prompt: Optional[str] = None
    include_visual_summaries: bool = False
    use_visual_context: bool = True
    include_logic: bool = False # 追加: 計算式や表示形式（単位）の抽出を有効にする
    dpi: int = 150
    max_file_size_mb: int = 50 # 追加: 巨大ファイルによるOOM防止のための制限 (MB)

class RagOptions(ExtractionOptions):
    """RAGチャンク生成時のオプション"""
    list_format: str = "kv"

class ExtractionResult(BaseModel):
    """LLMからの抽出結果を個別に保持するモデル"""
    model_config = ConfigDict(extra='allow')
    data: Optional[Union[List[Any], Dict[str, Any]]] = None
    _raw_json: Optional[str] = None
    error: Optional[str] = None

class SheetData(BaseModel):
    """シートごとの解析結果スキーマ（柔軟な構造を許容）"""
    model_config = ConfigDict(extra='allow')
    metadata: Optional[Dict[str, Any]] = Field(default_factory=dict)
    sections: Optional[List[Dict[str, Any]]] = Field(default_factory=list)
    data: Optional[Union[List[Any], Dict[str, Any]]] = Field(default_factory=list)
    cells: Optional[List[Dict[str, Any]]] = Field(default_factory=list) # 詳細セル情報
    html: Optional[str] = None # 抽出に使用したHTML
    media_map: Optional[Dict[str, List[Dict[str, Any]]]] = Field(default_factory=dict) # 追加: coord -> media_info
    errors: Optional[List[str]] = Field(default_factory=list)

class FullExtraction(BaseModel):
    """Excel全体の抽出データ"""
    model_config = ConfigDict(extra='allow')
    sheets: Dict[str, SheetData]
