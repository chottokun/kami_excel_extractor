# **ローカル完結型画像レンダリングとVLMを統合したExcelドキュメント構造化パイプラインの実現可能性および実装計画**

## **1\. エグゼクティブ・サマリーと実現可能性の総括**

企業内データにおける非構造化および半構造化データの大部分を占めるExcelスプレッドシートから、高度な精度で構造化データを抽出する自動化パイプラインの構築は、長年にわたりデータエンジニアリング分野における重大な課題であった。本調査の主たる目的は、視覚的言語モデル（VLM）以外の外部APIを一切利用せず、PythonとDocker Composeを基盤とした完全なローカル・コンテナ環境内で画像生成およびメタデータ解析を完結させ、最終的にGemini APIをはじめとする複数のLLM（Azure OpenAI、OpenAIなど）を自由に切り替え可能なアーキテクチャの実現可能性を検証し、その詳細な実装計画を立案することである。

調査および検証の結果、提示された制約条件下でのパイプライン構築は「極めて高い実現可能性」を有していると結論付けられる。このアーキテクチャの中核を成すのは、外部のOCRサービスやドキュメント解析API（Adobe PDF Extractなど）への依存を完全に排除し、ローカルのヘッドレスグラフィックエンジン（LibreOffice）を用いてExcelのグリッドを視覚的な画像情報（PNGまたはPDF）として高忠実度にレンダリングするプロセスである1。同時に、VLM単体では高密度の数値グリッドに対するOCRのハルシネーション（幻覚）や、数式・条件付き書式などの非可視メタデータの欠落が発生するリスクがあるため、Pythonライブラリ（openpyxl）を用いたローカルでのメタデータ抽出とコンテキスト圧縮を並行して実行するハイブリッド・アプローチが必須となる3。

さらに、LLMの切り替えをシームレスに行うためのゲートウェイ層としてLiteLLMをコンテナネットワーク内に配置することで、特定ベンダーへのロックインを完全に回避し、Gemini 2.0 Flashなどの最新モデルから、Azure OpenAIのGPT-4oやo1シリーズへのルーティングをアプリケーションコードの変更なしに実現可能である6。本報告書では、これら最新のコンテキスト最適化手法（SpreadsheetLLMの概念等）や並行処理の排他制御メカニズムを含めた、本番環境レベルの堅牢性を担保する実装計画を詳述する。

## **2\. アーキテクチャパラダイムと技術スタックの選定**

外部APIへの依存をVLMのエンドポイントのみに限定するという制約は、セキュリティ、データプライバシー、およびランニングコストの観点から非常に合理的である。近年、IBMのDoclingやLLMWhispererのようなレイアウト保持型のドキュメント変換ツールが登場しているが、これらはスキャンされた画像や複雑なExcelのレンダリングにおいて限界を示す場合がある9。したがって、ドキュメントの視覚的構造を人間が見るのと全く同じ状態でVLMに提示するためには、OSネイティブのレンダリングエンジンをコンテナ内にカプセル化するアプローチが最も確実である。

この要件を満たすため、本パイプラインは以下の要素技術によって構成される。これらの技術はすべてオープンソースであり、Docker Composeによって単一の仮想ネットワーク内でオーケストレーションされる。

| コンポーネント | 選定技術 | パイプラインにおける役割と選定理由 |
| :---- | :---- | :---- |
| **コンテナオーケストレーション** | Docker Compose | 複雑な依存関係を持つレンダリング環境と、軽量なAPIゲートウェイ、およびPythonオーケストレーターを分離し、再現性の高いローカル環境を構築する1。 |
| **LLMゲートウェイ / ルーター** | LiteLLM | アプリケーションとVLMの間に位置するミドルウェア。Gemini、Azure、OpenAIなど100以上のLLMプロバイダーのAPI差異を吸収し、単一のOpenAI互換フォーマットに標準化する6。 |
| **ローカル画像レンダリングエンジン** | LibreOffice (Headless) | .xlsxファイルを外部APIなしで高解像度のPNGまたはPDFに変換する。コンテナ内でXサーバーを必要とせずに動作し、完全な視覚的再現性を提供する1。 |
| **構造的メタデータ抽出** | openpyxl (Python) | レンダリング画像だけでは欠落するセルの背景色、罫線、結合状態などの論理フォーマットを解析・抽出する役割を担う3。 |
| **並行処理・プロセスマネージャー** | concurrent.futures / multiprocessing | Pythonの標準ライブラリを利用し、複数のExcelファイルの変換処理をマルチプロセスで実行する。シングルスレッド設計のLibreOfficeを安全に並行稼働させるための制御を行う15。 |

この構成における最大の利点は「関心の分離」である。ローカル環境は「視覚的状態の生成」と「構造データの抽出」という物理的・論理的タスクに専念し、VLMはそれらの入力を統合して「意味的理解と構造化JSONの生成」という推論タスクにのみ専念する。これにより、パイプライン全体の堅牢性と精度が飛躍的に向上する。

## **3\. VLMゲートウェイとしてのLiteLLMの導入と抽象化**

2025年以降、エンタープライズAIインフラストラクチャ市場の成熟に伴い、複数のLLMプロバイダー（Google、OpenAI、Anthropicなど）を併用するマルチモデル戦略が標準となりつつある6。しかし、各プロバイダーは独自のAPIスキーマ、認証方式、レート制限、ストリーミング仕様を持っており、これらをアプリケーション側で個別に実装・保守することは、深刻な「SDK地獄（SDK hell）」を引き起こす17。本計画の要件である「Gemini APIを初期利用し、将来的にAzure OpenAI等へ自由に切り替える」機能を実現するためには、LiteLLMをプロキシサーバーとしてDockerコンテナ内に構築することが最適解である。

### **プロバイダー非依存のルーティングメカニズム**

LiteLLMは、Pythonベースの抽象化レイヤーとして機能し、すべてのリクエストを標準的なOpenAIの ChatCompletions エンドポイントフォーマットで受け付け、それを背後の各プロバイダーが要求するネイティブフォーマットに動的かつ透過的に変換する7。

Docker Compose内でLiteLLMプロキシを起動する際、宣言的な config.yaml ファイルを用いてルーティングのルールを定義する。例えば、初期要件であるGemini API（Google AI Studio経由）を利用する場合、設定ファイルには gemini/ プレフィックスとターゲットモデル（例：gemini-2.5-pro または gemini-2.0-flash）を指定し、単純なAPIキーで認証を行う18。もし、エンタープライズのセキュリティ要件により、Google Cloud上のプライベート環境への移行が必要になれば、プレフィックスを vertex\_ai/ に変更し、GCPのサービスアカウント認証情報（JSON）を環境変数に渡すだけで、アプリケーションのコードを1行も変更することなくモデルのトラフィックを切り替えることが可能である19。

同様に、MicrosoftのAzure OpenAIに切り替える場合は、設定ファイルに azure/ プレフィックスを指定し、Azure固有の AZURE\_API\_BASE や AZURE\_API\_VERSION をマッピングする8。Pythonオーケストレーターは、一貫して http://litellm-proxy:4000/v1 というローカルのコンテナエンドポイントに対してのみリクエストを発行し続ける。

### **マルチモーダル・ペイロードの標準化とパラメータ変換**

画像やドキュメントをVLMに送信するマルチモーダル処理において、プロバイダー間のAPIの差異はさらに顕著になる。LiteLLMはこれらの複雑さを完全に隠蔽する。ローカルのLibreOfficeで生成されたExcelの画像はBase64エンコードされ、OpenAI標準の image\_url ペイロードとしてプロキシに送信される8。

Geminiモデルに対してリクエストがルーティングされた場合、LiteLLMはこのOpenAIフォーマットのマルチモーダルメッセージを自動的に解釈し、Google APIが要求する FileData オブジェクトやインラインデータパートに変換してAPIコールを実行する18。

さらに、LiteLLMはコスト最適化のための高度なパラメータ変換機能を提供する。例えば、OpenAIの reasoning\_effort（推論の深さ）パラメータは、Gemini環境においては自動的に thinking パラメータへとマッピングされる18。数千件のExcelファイルをバッチ処理する際など、モデルの深い推論が不要な単純な表抽出タスクにおいては、このパラメータを意図的に none に設定することで、APIコストを最大96%削減することが可能であるという検証結果も示されている18。

## **4\. ローカル画像レンダリングエンジン：LibreOffice Headlessの高度な制御**

外部APIを利用せずにExcelスプレッドシート（無限に広がる二次元のグリッド）を、VLMが解析可能な有限の画像フォーマット（PNGまたはPDF）に変換するプロセスは、本アーキテクチャにおいて最も技術的難易度が高いコンポーネントである。このタスクには、ローカルコンテナ内でGUIを起動せずに動作するLibreOfficeのヘッドレスモードを利用する1。

### **ヘッドレス変換の基本仕様とコンテナ環境の構築**

LibreOfficeは依存関係の多い巨大なアプリケーションであるため、ホストOSに直接インストールするのではなく、Dockerコンテナ内に隔離することが強く推奨される1。Dockerfileは ubuntu:22.04 等の軽量なベースイメージを使用し、libreoffice-calc および日本語・多言語対応のための包括的なフォント群（fonts-noto-cjk, fonts-liberation 等）をインストールする必要がある1。コンテナ内にExcelの元データで使用されているフォントが存在しない場合、LibreOfficeは代替フォントを強制適用するため、セル幅の計算が狂い、文字切れやレイアウトの崩壊が発生し、VLMの抽出精度を著しく低下させる原因となる1。

PythonからサブプロセスとしてLibreOfficeを呼び出す基本コマンドは以下のようになる。

soffice \--headless \--convert-to png \--outdir /app/output /app/input/data.xlsx 24

### **アーティファクトの除去（ヘッダー、フッター、余白の無効化）**

スプレッドシートを視覚情報に変換する際、重大な障害となるのが、レンダリングエンジンが自動的に付与するページングのアーティファクトである。LibreOffice Calcはデフォルトで、出力画像の上部に「Sheet 1」などのヘッダーを、下部に「Page 1」などのフッターを挿入し、さらに周囲に標準的な余白を確保する27。これらのアーティファクトは、貴重な画像のピクセル密度を浪費するだけでなく、VLMに対して「ページ番号がデータの一部である」という不要なノイズを与えてしまう28。

GUI上ではこれらをチェックボックスでオフにすることが可能であるが、ヘッドレス環境のコマンドラインからは直接制御するフラグが存在しない29。この問題に対する実現可能なプログラム的解決策として、以下の2つのアプローチが調査結果から導き出されている。

1. **UNO APIを介したマクロ制御:** PythonスクリプトからUniversal Network Objects (UNO) ブリッジを利用してバックグラウンドで起動しているLibreOfficeプロセスに接続し、ドキュメントの PageStyles を直接操作する手法である。アクティブシートのスタイルを取得し、.HeaderOn \= False および .FooterOn \= False プロパティを明示的に設定してからエクスポートを実行する30。  
2. **カスタムテンプレート（.ots）の適用:** より軽量かつ高速なアプローチとして、あらかじめGUIで余白をゼロにし、ヘッダーとフッターを完全に無効化した空のCalcテンプレートファイル（.ots）を作成しておく手法である27。変換コマンドのフィルタパラメータでこのテンプレートをベースとして強制適用することで、Python側の複雑なマクロ制御を省き、純粋なデータグリッドのみを画像として出力させることが可能となる。

### **並行処理時の競合回避とプロファイル分離**

パイプラインの運用において、複数のExcelファイルを同時に処理する必要が生じた場合、LibreOfficeのアーキテクチャ上の仕様が大きな障壁となる。LibreOfficeは本来デスクトップアプリケーションであるため、「単一のユーザープロファイルに対して同時に1つのアクティブインスタンスしか許可しない」という厳格な制約を持っている34。Pythonの concurrent.futures や multiprocessing を用いて複数の subprocess.run() を同時にトリガーした場合、後続のプロセスはハングアップするか、エラーを出力せずに空のファイルを生成して静かに失敗する（サイレントフェイラー）26。

この致命的な競合をローカル・コンテナ内で解決し、マルチコアの恩恵を最大限に受けるためには、Pythonオーケストレーターがプロセスをフォークするたびに、独立したエフェメラル（使い捨て）なユーザープロファイルディレクトリを動的に割り当てる必要がある36。

具体的には、変換コマンドに \-env:UserInstallation=file:///tmp/lo\_profile\_{uuid} という環境変数上書きパラメータを注入する36。これにより、各並行プロセスは互いに完全に独立した仮想的なLibreOfficeインスタンスとして振る舞い、リソースのロック競合を起こすことなく、安全かつ高速に複数ファイルの同時レンダリングを実行できるようになる36。

## **5\. openpyxlを用いたメタデータ抽出とコンテキスト最適化**

スプレッドシートの完全な高解像度画像を取得したとしても、それ単体をVLMに丸投げするアプローチには精度の限界がある。人間のアナリストは、背景色が黄色に塗られたセルを「注意すべき異常値」と認識し、太字の罫線で囲まれた領域を「独立したデータテーブル」と直感的に理解する。最新のVLM（Gemini 2.0 FlashやGPT-4o等）は極めて高い視覚的推論能力を持つが、数千のセルが密集する複雑なグリッドにおいて、細かな罫線の意味や、その数値がハードコードされたものか計算式（数式）による結果であるかを画像のみから完全に逆算することは不可能に近い5。

この情報欠落を補完し、VLMの推論を強力にガイドするためには、Pythonの openpyxl ライブラリを用いてExcelファイル（.xlsxの基底XML構造）を直接パースし、非可視の論理的メタデータやスタイル情報を抽出してプロンプトに結合するデュアル・モダリティ構成が不可欠である3。

### **意味的フォーマット情報の抽出メカニズム**

openpyxl を用いたPythonスクリプトは、ワークシート内の全セルをスキャンし、以下のような重要な構造的ヒントを抽出する。

* **セルの結合状態（Merge Cells）:** 結合されたセルは、階層的なヘッダー（複数の列を束ねる大項目など）を表すことが多く、表の親子関係を理解する上で最も重要な構造指標である。  
* **背景色と文字色:** cell.fill.start\_color.index 等を評価することで、RGB値やテーマカラーを取得する。特定の背景色（例：赤字や警告色）はデータのステータスや分類を示す14。  
* **罫線（Borders）のトポロジー:** cell.border 属性を検証することで、上下左右の罫線の有無や太さを取得する。これにより、ページ内に存在する複数の独立した表の境界線を論理的に特定できる41。  
* **データ型と数式の存在:** セルの値が日付型であるか、あるいは特定の数式（関数）によって導出された動的な値であるかを判定する。

### **SpreadsheetLLMの概念に基づくコンテキストの圧縮**

抽出したすべてのセルのフォーマット情報と値をそのままJSON配列にシリアライズしてLLMに送信すると、即座にトークン数の上限（コンテキストウィンドウ）を超過してしまうか、膨大なトークン消費による計算コストの増大を招く4。この問題を解決するため、本実装計画では2024年に提唱された「SpreadsheetLLM」および「SheetCompressor」の概念にインスパイアされたメタデータの圧縮アルゴリズムを導入する43。

この圧縮アルゴリズムの核心は「データフォーマット集約（Data-Format-Aware Aggregation）」と「構造的アンカーの抽出」である4。例えば、B2 から B100 までの列にすべて同一の日付フォーマットが適用され、特段の背景色が設定されていない場合、これら99個のセル情報を個別にJSON化するのではなく、{"range": "B2:B100", "datatype": "date", "style": "default\_column"} という単一のアンカー情報に圧縮する43。一方で、巨大なデータ領域から外れた特異なフォーマットのセル（例：太い赤枠で囲まれた D101 セルの合計値など）は、個別の座標情報としてJSONに保持する14。

この前処理により、数万セルに及ぶ巨大なシートのメタデータを、LLMのコンテキストを圧迫しない数十倍の圧縮率で軽量な構造マップ（JSON）へと変換することが可能となる4。この圧縮されたJSONマップは、視覚画像と合わせてVLMに渡される「地図」として機能する。

## **6\. マルチモーダルプロンプトエンジニアリングと構造化データ抽出**

ローカル環境で生成された「Excelの視覚的画像（PNG）」と「圧縮された構造的メタデータマップ（JSON）」という2つのアーティファクトが揃った後、これらを統合してLiteLLM経由でVLMへ送信するための高度なプロンプトエンジニアリングが必要となる。

近年のAIシステムにおいて、LLMに対するプロンプトは単なる自然言語の指示書きではなく、APIの仕様書と同等の「厳密な契約（Contract）」として機能しなければならない。このようなアプローチは「プロンプトの産業化（Prompt Industrialization）」と呼ばれており、自由記述のプロンプトがもたらす「曖昧さの税金（Ambiguity Tax）」（パースエラーや幻覚の修正にかかるコスト）を排除するために不可欠である45。

### **デュアル・モダリティを活用したプロンプト戦略**

VLMに対して精度の高い抽出を行わせるための最適なプロンプト構成は、段階的な推論を強要する構造を取る39。まず、システムプロンプトでLLMに「複雑なグリッドデータから構造を抽出する専門のデータアナリスト」というペルソナを与える45。

次に、画像とテキストデータの相関関係を明示的に指示する。VLMは画像全体を均等に解釈しようとするため、注目すべき領域（Region of Interest）をテキスト側からガイドする必要がある39。例えば、プロンプト内で次のように指示する。「提供されたJSONメタデータマップを参照してください。背景色が \#FFFF00（黄色）で示されたセル座標は、抽出すべき重要KPIを示しています。画像内でこれらの座標に該当する視覚的領域を特定し、その左側に隣接する行ヘッダーのテキストと関連付けて抽出してください」14。

この手法により、VLMは「画像を単にOCRする」のではなく、「論理的な座標マップ（JSON）と視覚的なピクセル情報（画像）を相互検証しながらデータを抽出する」という高度な推論タスクを実行するようになり、抽出精度が劇的に向上する39。

### **Constrained Decoding（制約付きデコーディング）によるJSONスキーマの強制**

構造化データの抽出において、過去のLLMは「こちらが抽出したJSONです：」といった不要な会話文を付与したり、閉じカッコを忘れたりすることで、後続のシステムパイプラインを破壊することが頻発していた45。しかし、本計画で利用するGemini 1.5 Pro/2.0 Flashや、切り替え先となるAzure OpenAIのGPT-4o等の最新モデルは、APIレベルでの「JSONモード」および「制約付きデコーディング（Constrained Decoding）」を強力にサポートしている22。

Pythonアプリケーション側でPydanticを用いて、出力として期待する厳密なデータスキーマ（クラス構造）を定義する。LiteLLMを通じてこのスキーマをAPIリクエストの response\_format としてプロバイダーに送信することで、モデルは生成プロセスのトークン選択レベルで制約を受ける。つまり、指定されたJSONスキーマに違反するトークンを物理的に生成できなくなるため、99%以上の確率で構文的に完璧なJSON出力を保証することができる45。これにより、出力結果に対する正規表現を用いた脆い後処理スクリプトやリトライ処理をコードベースから一掃することが可能となる。

## **7\. コンテナオーケストレーションと並行処理の実装**

ここまでに定義されたすべてのコンポーネントを、安全かつスケール可能なローカル・インフラストラクチャとしてデプロイするため、Docker Composeを用いたアーキテクチャを実装する。ネットワークのトポロジーは、大きく2つのサービスレイヤーに分割される。

### **サービス定義とDocker Compose構成**

1. **litellm-proxy サービス:** 外部のVLMへのアウトバウンド通信を全て担う軽量なAPIゲートウェイ。公式の ghcr.io/berriai/litellm イメージを使用し、内部ポート 4000 を開放する7。APIキー（Gemini APIキー、Azure認証情報など）はセキュアな環境変数から注入され、config.yaml の設定に基づいてロードバランシングやフェイルオーバーの管理を行う7。  
2. **pipeline-worker サービス:** データ処理の心臓部となるカスタムコンテナ。ベースイメージには ubuntu:22.04 などのフル機能OS環境を採用し、LibreOffice、必要なフォント群、Python 3.10以上、および openpyxl, openai, pydantic 等のライブラリをインストールする1。外部のインターネットへのアクセスはLiteLLMコンテナへのAPIコールのみに制限することで、高度なセキュリティ環境を担保できる。

### **Pythonによるマルチプロセス・オーケストレーション**

pipeline-worker コンテナ内で稼働するPythonメインプロセスは、継続的に特定のローカルディレクトリ（ボリュームマウントされた入力フォルダ）を監視し、新しいExcelファイルが配置されると処理を開始する。

大量のドキュメントを迅速に処理するため、Pythonの concurrent.futures.ProcessPoolExecutor または multiprocessing.Pool を利用してワークロードを分散する15。各ワーカースレッドにファイルが割り当てられると、前述の動的プロファイル生成技術（-env:UserInstallation）を用いてLibreOfficeサブプロセスを安全に起動し36、並行して openpyxl によるメタデータのパースと圧縮を実行する。

両方の前処理（画像のBase64化とJSONマップの生成）が完了したスレッドは、ローカルの http://litellm-proxy:4000/v1/chat/completions に対して非同期のRESTリクエストを発行する7。これにより、CPUバウンドなローカルレンダリング処理と、I/OバウンドなLLM API待ち時間が効率的に多重化され、コンテナのリソースを極限まで引き出した高速なパイプライン処理が実現する16。

Azure OpenAI環境へ移行し、GPT-4oなどのモデルでAzure固有のVision機能（OCR、Grounding）を使用したい場合でも、LiteLLMは透過的に enhancements パラメータや dataSources（Azure Computer Visionエンドポイント）を中継するため、このPythonの非同期処理アーキテクチャに変更を加える必要は一切ない8。

## **8\. 詳細実装ロードマップ**

本アーキテクチャの確実な具現化に向け、以下の4フェーズからなる実装計画を立案する。外部APIへの依存がないローカル処理の安定化を最優先とし、その後に確率論的なLLMの統合を行う。

### **フェーズ 1: インフラストラクチャの初期化とゲートウェイ構築**

1. **カスタムDockerfileの策定:** UbuntuベースにLibreOffice Headless、多言語対応のNoto/Liberationフォント群、Python実行環境を統合した堅牢なイメージを構築する1。  
2. **LiteLLMルーターの構成:** litellm-config.yaml を作成し、主系として gemini/gemini-2.0-flash へのルーティング、副系として azure/gpt-4o へのフェイルオーバー経路を定義する8。  
3. **コンテナネットワークの起動:** docker-compose.yml を記述し、各コンテナ間の通信を確立し、環境変数によるセキュアなシークレット管理を実装する7。

### **フェーズ 2: ローカル前処理エンジンとレンダリングの確立**

1. **メタデータ抽出・圧縮モジュールの開発:** openpyxl を駆使し、セルの結合状態、背景色、罫線情報を取得するPythonスクリプトを開発する3。連続する同一フォーマットのセルを単一の構造アンカーに集約する圧縮アルゴリズム（SpreadsheetLLMベース）を実装し、軽量なJSONを生成する4。  
2. **アーティファクト除去テンプレートの作成:** 余白ゼロ、ヘッダー・フッター無効の専用Calcテンプレート（.ots）を作成し、変換コマンドに組み込む30。  
3. **安全なプロセスラッパーの開発:** UserInstallation パラメータを動的に付与し、LibreOfficeプロセスを隔離実行するPythonラッパー関数を実装し、サイレントフェイラーを根絶する26。

### **フェーズ 3: VLMの統合とプロンプトの実装**

1. **マルチモーダルペイロードフォーマッター:** 抽出されたメタデータJSONと、生成されたBase64画像を結合し、OpenAI互換のメッセージ配列（type: image\_url および type: text の混在構造）を生成するモジュールを実装する8。  
2. **スキーマ定義と制約の適用:** Pydanticを用いて目標とする抽出データのスキーマを定義し、APIリクエストに付与することで、GeminiおよびAzure OpenAIに対して厳格なJSON出力を強制する22。  
3. **プロンプトのチューニング:** 視覚情報とメタデータマップの対応付けを指示する高精度なシステムプロンプトを構築し、ハルシネーションを最小化する39。

### **フェーズ 4: テスト、チューニング、そして拡張**

1. **並行処理負荷テスト:** 大量のExcelドキュメントを同時に投入し、ProcessPoolExecutor によるプロセス割り当てとLibreOfficeの安定稼働、およびメモリリークの有無を検証する37。  
2. **モデル間ベンチマーク:** LiteLLMの切り替え機能を利用し、Gemini 1.5 Pro、Gemini 2.0 Flash、Azure GPT-4o間で、抽出精度、レイテンシ、トークン消費量を比較・評価する6。  
3. **コスト最適化の自動化:** 抽出対象の複雑さに応じて、LiteLLM側で推論パラメータ（reasoning\_effort のオンオフなど）を動的に切り替え、APIコストとパフォーマンスの最適なバランス地点を模索する18。

以上の計画により、最新のコンテナ技術、高度なローカルレンダリング制御、そして抽象化されたLLMゲートウェイを組み合わせることで、強固なセキュリティを担保しつつ、将来のモデル進化に追従可能な最高水準のExcelドキュメント構造化パイプラインが実現可能であると確信される。

#### **引用文献**

1. How to Run LibreOffice in Docker for Document Conversion \- OneUptime, 3月 1, 2026にアクセス、 [https://oneuptime.com/blog/post/2026-02-08-how-to-run-libreoffice-in-docker-for-document-conversion/view](https://oneuptime.com/blog/post/2026-02-08-how-to-run-libreoffice-in-docker-for-document-conversion/view)  
2. Building a Multimodal RAG That Responds with Text, Images, and Tables from Sources, 3月 1, 2026にアクセス、 [https://towardsdatascience.com/building-a-multimodal-rag-with-text-images-tables-from-sources-in-response/](https://towardsdatascience.com/building-a-multimodal-rag-with-text-images-tables-from-sources-in-response/)  
3. Formatting Cells using openpyxl in Python \- GeeksforGeeks, 3月 1, 2026にアクセス、 [https://www.geeksforgeeks.org/python/formatting-cells-using-openpyxl-in-python/](https://www.geeksforgeeks.org/python/formatting-cells-using-openpyxl-in-python/)  
4. SpreadsheetLLM: Encoding Spreadsheets for Large Language Models \- Microsoft Research, 3月 1, 2026にアクセス、 [https://www.microsoft.com/en-us/research/publication/encoding-spreadsheets-for-large-language-models/](https://www.microsoft.com/en-us/research/publication/encoding-spreadsheets-for-large-language-models/)  
5. Spreadsheet QnA with LLMs: Finding the Optimal Representation Format and the Strategy \- Part-1 \- Quantiphi, 3月 1, 2026にアクセス、 [https://quantiphi.com/blog/spreadsheet-qna-with-llms-finding-the-optimal-representation-format-and-the-strategy-part-1/](https://quantiphi.com/blog/spreadsheet-qna-with-llms-finding-the-optimal-representation-format-and-the-strategy-part-1/)  
6. Top 5 LiteLLM Alternatives in 2025 \- DEV Community, 3月 1, 2026にアクセス、 [https://dev.to/debmckinney/top-5-litellm-alternatives-in-2025-1pki](https://dev.to/debmckinney/top-5-litellm-alternatives-in-2025-1pki)  
7. Getting Started \- LiteLLM Docs, 3月 1, 2026にアクセス、 [https://docs.litellm.ai/docs/](https://docs.litellm.ai/docs/)  
8. Azure OpenAI | liteLLM, 3月 1, 2026にアクセス、 [https://docs.litellm.ai/docs/providers/azure/](https://docs.litellm.ai/docs/providers/azure/)  
9. docling-project/docling: Get your documents ready for gen AI \- GitHub, 3月 1, 2026にアクセス、 [https://github.com/docling-project/docling](https://github.com/docling-project/docling)  
10. Docling vs. LLMWhisperer: Best Docling Alternative in 2026 \- Unstract, 3月 1, 2026にアクセス、 [https://unstract.com/blog/docling-alternative/](https://unstract.com/blog/docling-alternative/)  
11. Docling is a new library from IBM that efficiently parses PDF, DOCX, and PPTX and exports them to Markdown and JSON. : r/LocalLLaMA \- Reddit, 3月 1, 2026にアクセス、 [https://www.reddit.com/r/LocalLLaMA/comments/1ghbmoq/docling\_is\_a\_new\_library\_from\_ibm\_that/](https://www.reddit.com/r/LocalLLaMA/comments/1ghbmoq/docling_is_a_new_library_from_ibm_that/)  
12. LibreOffice+Docker+Express/Lambda: Convert Office To PDF. Serverless. For Free\!, 3月 1, 2026にアクセス、 [https://levelup.gitconnected.com/libreoffice-docker-express-lambda-convert-office-to-pdf-serverless-for-free-8781bc2f0c55](https://levelup.gitconnected.com/libreoffice-docker-express-lambda-convert-office-to-pdf-serverless-for-free-8781bc2f0c55)  
13. Working with styles — openpyxl 3.1.4 documentation, 3月 1, 2026にアクセス、 [https://openpyxl.readthedocs.io/en/3.1/styles.html](https://openpyxl.readthedocs.io/en/3.1/styles.html)  
14. Extracting Color Codes from Excel Files in Python with OpenPyXL and XLRD \- Medium, 3月 1, 2026にアクセス、 [https://medium.com/@waseem3378/extracting-color-codes-from-excel-files-in-python-with-openpyxl-and-xlrd-18c9bca40d94](https://medium.com/@waseem3378/extracting-color-codes-from-excel-files-in-python-with-openpyxl-and-xlrd-18c9bca40d94)  
15. multiprocessing — Process-based parallelism — Python 3.14.3 documentation, 3月 1, 2026にアクセス、 [https://docs.python.org/3/library/multiprocessing.html](https://docs.python.org/3/library/multiprocessing.html)  
16. How to handle multiple subprocesses simultaneously \- Discussions on Python.org, 3月 1, 2026にアクセス、 [https://discuss.python.org/t/how-to-handle-multiple-subprocesses-simultaneously/59024](https://discuss.python.org/t/how-to-handle-multiple-subprocesses-simultaneously/59024)  
17. SDK hell with multiple LLM providers? Compared LangChain, LiteLLM, and any-llm \- Reddit, 3月 1, 2026にアクセス、 [https://www.reddit.com/r/LLMDevs/comments/1njelu2/sdk\_hell\_with\_multiple\_llm\_providers\_compared/](https://www.reddit.com/r/LLMDevs/comments/1njelu2/sdk_hell_with_multiple_llm_providers_compared/)  
18. Gemini \- Google AI Studio | liteLLM, 3月 1, 2026にアクセス、 [https://docs.litellm.ai/docs/providers/gemini](https://docs.litellm.ai/docs/providers/gemini)  
19. VertexAI \[Gemini\] \- LiteLLM Docs, 3月 1, 2026にアクセス、 [https://docs.litellm.ai/docs/providers/vertex](https://docs.litellm.ai/docs/providers/vertex)  
20. OpenAI \- LiteLLM Docs, 3月 1, 2026にアクセス、 [https://docs.litellm.ai/docs/providers/openai](https://docs.litellm.ai/docs/providers/openai)  
21. Process images, video, audio, and text with Gemini 1.5 Pro | Generative AI on Vertex AI, 3月 1, 2026にアクセス、 [https://docs.cloud.google.com/vertex-ai/generative-ai/docs/samples/generativeaionvertexai-gemini-all-modalities](https://docs.cloud.google.com/vertex-ai/generative-ai/docs/samples/generativeaionvertexai-gemini-all-modalities)  
22. Multi-Modal LLM using Google's Gemini model for image understanding and build Retrieval Augmented Generation with LlamaIndex, 3月 1, 2026にアクセス、 [https://developers.llamaindex.ai/python/examples/multi\_modal/gemini/](https://developers.llamaindex.ai/python/examples/multi_modal/gemini/)  
23. node.js \- How can i add libreoffice to Dockerfile \- Stack Overflow, 3月 1, 2026にアクセス、 [https://stackoverflow.com/questions/75109416/how-can-i-add-libreoffice-to-dockerfile](https://stackoverflow.com/questions/75109416/how-can-i-add-libreoffice-to-dockerfile)  
24. Use Soffice to convert from doc to png images \- English \- Ask LibreOffice, 3月 1, 2026にアクセス、 [https://ask.libreoffice.org/t/use-soffice-to-convert-from-doc-to-png-images/41621](https://ask.libreoffice.org/t/use-soffice-to-convert-from-doc-to-png-images/41621)  
25. Command line convert to png issue \- English \- Ask LibreOffice, 3月 1, 2026にアクセス、 [https://ask.libreoffice.org/t/command-line-convert-to-png-issue/5686](https://ask.libreoffice.org/t/command-line-convert-to-png-issue/5686)  
26. Serving Concurrent Requests for LibreOffice Service \- jdhao, 3月 1, 2026にアクセス、 [https://jdhao.github.io/2021/06/11/libreoffice\_concurrent\_requests/](https://jdhao.github.io/2021/06/11/libreoffice_concurrent_requests/)  
27. How to remove header and footer in Calc by default? \- English \- Ask LibreOffice, 3月 1, 2026にアクセス、 [https://ask.libreoffice.org/t/how-to-remove-header-and-footer-in-calc-by-default/28239](https://ask.libreoffice.org/t/how-to-remove-header-and-footer-in-calc-by-default/28239)  
28. Export spreadsheet to PNG without header/footer \- libreoffice \- Reddit, 3月 1, 2026にアクセス、 [https://www.reddit.com/r/libreoffice/comments/1ik7gbc/export\_spreadsheet\_to\_png\_without\_headerfooter/](https://www.reddit.com/r/libreoffice/comments/1ik7gbc/export_spreadsheet_to_png_without_headerfooter/)  
29. How do I delete the header and footer to maximize print space on page? \- Ask LibreOffice, 3月 1, 2026にアクセス、 [https://ask.libreoffice.org/t/how-do-i-delete-the-header-and-footer-to-maximize-print-space-on-page/63946](https://ask.libreoffice.org/t/how-do-i-delete-the-header-and-footer-to-maximize-print-space-on-page/63946)  
30. Print documents without footer information \- English \- Ask LibreOffice, 3月 1, 2026にアクセス、 [https://ask.libreoffice.org/t/print-documents-without-footer-information/81935](https://ask.libreoffice.org/t/print-documents-without-footer-information/81935)  
31. LibreOffice Calc macro \- How to remove header and footer in given sheet? \- Stack Overflow, 3月 1, 2026にアクセス、 [https://stackoverflow.com/questions/30213384/libreoffice-calc-macro-how-to-remove-header-and-footer-in-given-sheet](https://stackoverflow.com/questions/30213384/libreoffice-calc-macro-how-to-remove-header-and-footer-in-given-sheet)  
32. File Conversion Filters Tables \- LibreOffice Help, 3月 1, 2026にアクセス、 [https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html](https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html)  
33. Creating and Changing Default and Custom Templates \- LibreOffice Help, 3月 1, 2026にアクセス、 [https://help.libreoffice.org/latest/en-US/text/shared/guide/standard\_template.html](https://help.libreoffice.org/latest/en-US/text/shared/guide/standard_template.html)  
34. Is LibreOffice (headless) safe to use on a web server? \- Stack Overflow, 3月 1, 2026にアクセス、 [https://stackoverflow.com/questions/55070766/is-libreoffice-headless-safe-to-use-on-a-web-server](https://stackoverflow.com/questions/55070766/is-libreoffice-headless-safe-to-use-on-a-web-server)  
35. executing commands in containers from within a container using docker-compose up vs docker-compose run \- Stack Overflow, 3月 1, 2026にアクセス、 [https://stackoverflow.com/questions/50134704/executing-commands-in-containers-from-within-a-container-using-docker-compose-up](https://stackoverflow.com/questions/50134704/executing-commands-in-containers-from-within-a-container-using-docker-compose-up)  
36. Command \`libreoffice \--headless \--convert-to pdf test.docx \--outdir /pdf\` is not working \[closed\] \- Stack Overflow, 3月 1, 2026にアクセス、 [https://stackoverflow.com/questions/30349542/command-libreoffice-headless-convert-to-pdf-test-docx-outdir-pdf-is-not](https://stackoverflow.com/questions/30349542/command-libreoffice-headless-convert-to-pdf-test-docx-outdir-pdf-is-not)  
37. Running several instances of headless libreoffice: where's race?, 3月 1, 2026にアクセス、 [https://ask.libreoffice.org/t/running-several-instances-of-headless-libreoffice-wheres-race/104816](https://ask.libreoffice.org/t/running-several-instances-of-headless-libreoffice-wheres-race/104816)  
38. subprocess not working in docker container \- Stack Overflow, 3月 1, 2026にアクセス、 [https://stackoverflow.com/questions/49439955/subprocess-not-working-in-docker-container](https://stackoverflow.com/questions/49439955/subprocess-not-working-in-docker-container)  
39. multimodal prompting : r/PromptEngineering \- Reddit, 3月 1, 2026にアクセス、 [https://www.reddit.com/r/PromptEngineering/comments/1jg6c74/multimodal\_prompting/](https://www.reddit.com/r/PromptEngineering/comments/1jg6c74/multimodal_prompting/)  
40. OpenPyXL: Color Cells Without the Tedious Clicking \- YouTube, 3月 1, 2026にアクセス、 [https://www.youtube.com/watch?v=0AlQtxFqv54](https://www.youtube.com/watch?v=0AlQtxFqv54)  
41. python \- OpenPyXL \- How to query cell borders? \- Stack Overflow, 3月 1, 2026にアクセス、 [https://stackoverflow.com/questions/51900841/openpyxl-how-to-query-cell-borders](https://stackoverflow.com/questions/51900841/openpyxl-how-to-query-cell-borders)  
42. openpyxl.styles.borders module \- Read the Docs, 3月 1, 2026にアクセス、 [https://openpyxl.readthedocs.io/en/3.1/api/openpyxl.styles.borders.html](https://openpyxl.readthedocs.io/en/3.1/api/openpyxl.styles.borders.html)  
43. How to Fit Massive Excel Files into LLMs: The Spreadsheet Compression Playbook | by Denis Urayev | Medium, 3月 1, 2026にアクセス、 [https://medium.com/@denisuraev/how-to-fit-massive-excel-files-into-llms-the-spreadsheet-compression-playbook-051173d0331d](https://medium.com/@denisuraev/how-to-fit-massive-excel-files-into-llms-the-spreadsheet-compression-playbook-051173d0331d)  
44. SpreadsheetLLM: Encoding Spreadsheets for Large Language Models \- arXiv, 3月 1, 2026にアクセス、 [https://arxiv.org/html/2407.09025v1](https://arxiv.org/html/2407.09025v1)  
45. Structured Prompting with JSON: The Engineering Path to Reliable LLMs | by vishal dutt, 3月 1, 2026にアクセス、 [https://medium.com/@vishal.dutt.data.architect/structured-prompting-with-json-the-engineering-path-to-reliable-llms-2c0cb1b767cf](https://medium.com/@vishal.dutt.data.architect/structured-prompting-with-json-the-engineering-path-to-reliable-llms-2c0cb1b767cf)  
46. Using advanced prompt engineering techniques to create a data analyst \- Reddit, 3月 1, 2026にアクセス、 [https://www.reddit.com/r/PromptEngineering/comments/1ebrdeq/using\_advanced\_prompt\_engineering\_techniques\_to/](https://www.reddit.com/r/PromptEngineering/comments/1ebrdeq/using_advanced_prompt_engineering_techniques_to/)  
47. Multimodal RAG with Vision: From Experimentation to Implementation \- ISE Developer Blog, 3月 1, 2026にアクセス、 [https://devblogs.microsoft.com/ise/multimodal-rag-with-vision/](https://devblogs.microsoft.com/ise/multimodal-rag-with-vision/)  
48. The Guide to JSON: Parsing and Formatting Data Structures with Gemini, 3月 1, 2026にアクセス、 [https://geminiprompt.id/blog/the-guide-to-json-parsing-and-formatting-data-structures-with-gemini](https://geminiprompt.id/blog/the-guide-to-json-parsing-and-formatting-data-structures-with-gemini)  
49. Centralizing Multiple AI Services with LiteLLM Proxy | by Robert McDermott \- Medium, 3月 1, 2026にアクセス、 [https://robert-mcdermott.medium.com/centralizing-multi-vendor-llm-services-with-litellm-9874563f3062](https://robert-mcdermott.medium.com/centralizing-multi-vendor-llm-services-with-litellm-9874563f3062)  
50. docker-libreoffice-headless/Dockerfile at master \- GitHub, 3月 1, 2026にアクセス、 [https://github.com/ipunkt/docker-libreoffice-headless/blob/master/Dockerfile](https://github.com/ipunkt/docker-libreoffice-headless/blob/master/Dockerfile)  
51. Parallel programming in Python: multiprocessing (part 1\) – PDC Blog \- KTH, 3月 1, 2026にアクセス、 [https://www.kth.se/blogs/pdc/2019/02/parallel-programming-in-python-multiprocessing-part-1/](https://www.kth.se/blogs/pdc/2019/02/parallel-programming-in-python-multiprocessing-part-1/)  
52. Running external commands in parallel \- Python discussion forum, 3月 1, 2026にアクセス、 [https://discuss.python.org/t/running-external-commands-in-parallel/17912](https://discuss.python.org/t/running-external-commands-in-parallel/17912)  
53. Gemini 3 Flash Preview \- API, Providers, Stats \- OpenRouter, 3月 1, 2026にアクセス、 [https://openrouter.ai/google/gemini-3-flash-preview](https://openrouter.ai/google/gemini-3-flash-preview)