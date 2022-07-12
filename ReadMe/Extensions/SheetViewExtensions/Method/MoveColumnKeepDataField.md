# SheetViewExtensions.MoveColumnKeepDataField メソッド
# 定義
名前空間: Metroit.Win.GcSpread.CodeDesign

DataField プロパティを保持したまま列の移動を行います。

# MoveColumnKeepDataField(int, int, int)
DataField プロパティを保持したまま列の移動を行います。  
SheetView.MoveColumn(int, int, int, true) によって列の移動が行われます。
```cs
public static List<ColumnMoveResult> MoveColumnKeepDataField(int fromIndex, int toIndex, int columnCount)
```

### パラメーター
fromIndexx  
移動元の列インデックス。  
toIndex  
移動先の列インデックス。  
columnCount  
移動する列数。

### 戻り値
移動によって影響を受けた列の移動情報を返却します。  
OriginalColumnIndex, BeforeColumnIndex プロパティは常に同一のものを返却します。
