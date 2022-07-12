# SheetViewExtensions
GrapeCity SPREAD for Windows Forms の SheetView オブジェクトのレイアウトをjsonデータからバインドする拡張メソッドを提供します。

# 定義
名前空間: Metroit.Win.GcSpread.CodeDesign

# コンストラクタ―
なし

# フィールド
なし

# プロパティ
なし

# 拡張メソッド
|メソッド|説明|
|----|----|
|[BindJsonLayout(string, Func<SheetView, Root, IList<ColumnMoveResult>>)](Method/BindJsonLayout.md)|指定したJSONファイルでシートをバインドします。|
|[BindJsonLayout(SheetViewDefinitions, Func<SheetView, Root, IList<ColumnMoveResult>>)](Method/BindJsonLayout.md)|指定したJSONオブジェクトでシートをバインドします。|
|[CreateSheetDefinitions(Func<FarPoint.Win.Spread.Column, Dictionary<string, object>>)](Method/CreateSheetDefinitions.md)|シート情報からJSONデータオブジェクトを生成します。|
|[CreateSpecialColumnHeaderDefinitions()](Method/CreateSpecialColumnHeaderDefinitions.md)|シート情報から列ヘッダーの特定のみのJSONデータオブジェクトを生成します。|
|[CreateSpecialColumnDefinitions()](Method/CreateSpecialColumnDefinitions.md)|シート情報から列の特定のみのJSONデータオブジェクトを生成します。|
|[MoveColumnKeepDataField(int, int, int)](Method/MoveColumnKeepDataField.md)|DataField プロパティを保持したまま列の移動を行います。|
