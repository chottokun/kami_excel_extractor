# **汎用構造メタデータ抽出と適応型マッピング**

「標準的な表形式」「方眼紙エクセル」「報告書形式」が混在する環境に対応するため、Excelを「視覚的座標を持つ情報の集合」として解析し、VLMが解釈可能な「構造マップ」を生成する。

## **1. 全方位型メタデータ抽出 (Universal Extraction)**

`openpyxl` を用い、特定のレイアウトを仮定せずに以下の情報を全セルから抽出する。

- **絶対座標マップ:** 全セルのテキストと、その A1 形式の座標。
- **結合セル・トポロジー:** 結合された範囲（例：`B2:F3`）を一つの「論理ユニット」として定義。
- **視覚的ヒント (Visual Hints):**
    - **罫線 (Borders):** 上下左右の線の有無と太さ。表の境界や入力ボックスの特定に利用。
    - **背景色 (Fill):** 特定の項目（見出しや重要項目）の強調を特定。
- **データ型と数式:** 画像からは判別不能な「計算式」や「正確な数値型」を保持。

## **2. 適応型コンテキスト圧縮 (Adaptive Compression)**

SpreadsheetLLM の概念を拡張し、レイアウトの種類に応じて情報を整理する。

### **A. 表形式領域の検出**
連続して同じデータ型やスタイルが並ぶ領域を `TableRange` として集約し、VLMに「ここは単純な表である」と教える。

### **B. 自由形式（方眼紙・報告書）の構造化**
「ラベル（項目名）」と「バリュー（入力値）」のペアを特定するためのヒントを生成。
- 罫線で囲まれた孤立したセル群を `InputField` としてマーク。
- 結合セルを `InformationBlock` として定義。

## **3. 抽出エンジン（Python）の強化案**

```python
def generate_universal_map(ws):
    cells_map = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None and not cell.border: continue # 空セルはスキップして圧縮
            
            cell_info = {
                "coord": cell.coordinate,
                "val": str(cell.value) if cell.value else "",
                "style": {
                    "bg": cell.fill.start_color.index if cell.fill else None,
                    "borders": get_border_info(cell),
                    "is_merged": cell.coordinate in ws.merged_cells
                }
            }
            cells_map.append(cell_info)
    return cells_map
```

## **4. VLMへの「地図」としての提供**

生成されたメタデータは、画像を読み解くための「透明なレイヤー」として機能する。
VLMは画像から「だいたいの位置」を特定し、メタデータJSONから「正確な値」を引くという動作を行う。
これにより、**方眼紙形式での位置関係の誤認や、表形式での行ズレを物理的に防止する。**
