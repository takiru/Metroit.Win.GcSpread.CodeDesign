# SheetViewExtensions.BindJsonLayout メソッド
# 定義
名前空間: Metroit.Win.GcSpread.CodeDesign

指定したJSONでシートをバインドします。

# BindJsonLayout(string, Func<Cell, object>, Func<SheetView, Root, IList<ColumnMoveResult>>)
指定したJSONファイルでシートをバインドします。
```cs
public static void BindJsonLayout(string path,
            Func<Cell, object> columnHeaderCellTag,
            Func<SheetView, Root, IList<ColumnMoveResult>> intercepter = null)
```

### パラメーター
path  
jsonファイルのフルパス。  
columnHeaderCellTag  
列ヘッダーセルの Tag プロパティを任意の値に設定するデリゲート。  
intercepter  
バインド途中に差し込む任意のコードデリゲート。

### 戻り値
なし。

### 動作
動作の詳細は BindJsonLayout(SheetViewDefinitions, Func<Cell, object>, Func<SheetView, Root, IList<ColumnMoveResult>>) と同一です。


# BindJsonLayout(SheetViewDefinitions, Func<Cell, object>, Func<SheetView, Root, IList<ColumnMoveResult>>)
指定したJSONファイルでシートをバインドします。
```cs
public static void BindJsonLayout(SheetViewDefinitions jsonObj,
            Func<Cell, object> columnHeaderCellTag,
            Func<SheetView, Root, IList<ColumnMoveResult>> intercepter = null)
```

### パラメーター
jsonObj  
SheetViewDefinitions オブジェクト。  
columnHeaderCellTag  
列ヘッダーセルの Tag プロパティを任意の値に設定するデリゲート。  
intercepter  
バインド途中に差し込む任意のコードデリゲート。

### 戻り値
なし。

### SheetView の設定
SheetView の下記のプロパティはすべて false である必要があります。
  - AutoGenerateColumns
  - DataAutoCellTypes
  - DataAutoSizeColumns
  - DataAutoHeadings

### バインドされる情報
columnHeaderCellTag が設定されていない場合、列ヘッダーは、下記の情報が追加で設定されます。
|プロパティ|値|
|----------|--|
|SheetView.ColumnHeader.Cells[row, column].Tag|"Header_{row}_{column}"|

Tag プロパティの値は、列移動を伴ったとしても変化しません。  
これにより、Tag プロパティを調べることで、特定の列ヘッダーセルを取得することができます。  

列に DataField をバインドする場合、下記の情報が設定されます。
|プロパティ|値|
|----------|--|
|SheetView.Columns[column].DataField|jsonObj.Columns[column] で指定されている DataField 値|

DataField は、 SheetView.BindDataColumn(column, dataField) によって設定されます。  
複数列に同一の DataField が設定される場合、GrapeCity SPREAD for Windows Forms の仕様により、列情報を一意に定めることができません。  
この場合、 DataField プロパティを調べたあと、SheetView.ColumnHeader.Cells[row, column].Tag などの情報から目的の列を求める必要があります。  

### columnHeaderCellTag
列ヘッダーの Cell.Tag プロパティに任意の情報を設定したい場合に指定します。

### interceptor
意図的に jsonデータを超越して制御を行いたい場合に指定します。  
メソッドが実行された時の SheetView に存在する列の順序は jsonデータの順に準拠します。  
列ヘッダーのタイトルの変更、列の移動、アプリケーション固有の値制御などを行うことができます。  
列の移動を行う場合、列の移動は必ず MoveColumnKeepDataField() で行う必要があり、最終的な列の移動結果を返却する必要があります。

例
```cs
fpSpread1.ActiveSheet.BindJsonLayout(root,
    (sheetView, root2) =>
    {
        // ここで、Meestroテーブルのユーザー単位の列位置、タイトルなどで書き換える。
        // Columns[].Options プロパティを Tag に入れるなどもここで行う。
        // 列位置は下記コードによって行い、列移動結果を返却する必要がある。
        // 列移動がない場合は戻り値は null でよい。

        var result = new List<ColumnMoveResult>();

        // 1列目1行目ヘッダーのタイトルを変更する。
        sheetView.ColumnHeader.Cells[0, 0].Value = "値変更";

        // 列移動情報と仮定
        var columnMoves = new List<(int beforeIndex, int afterIndex, int columnCount)>()
        {
            (0, 3, 1),
            (1, 0, 2),
            (0, 5, 2)
        };

        foreach (var columnMove in columnMoves)
        {
            // 列移動をする
            // DataField プロパティを保持する必要があるため、MoveColumnKeepDataField() の実行が必要。
            var cmrs = sheetView.MoveColumnKeepDataField(columnMove.beforeIndex, columnMove.afterIndex, columnMove.columnCount);

            // 既に移動しているものがあったらオリジナルのインデックスを保持して書き換えて、resultからは削除する
            for (var i = 0; i < cmrs.Count; i++)
            {
                var exists = result.Where(x => x.AfterColumnIndex == cmrs[i].BeforeColumnIndex).FirstOrDefault();
                if (exists != null)
                {
                    cmrs[i] = new ColumnMoveResult(exists.OriginalColumnIndex, cmrs[i].BeforeColumnIndex, cmrs[i].AfterColumnIndex, cmrs[i].DataField);
                    result.Remove(exists);
                }
            }
            // 移動した分を把握する
            foreach (var cmr in cmrs)
            {
                result.Add(cmr);
            }
        }

        return result;
    });
```

### 例外
InvalidOperationException  
SheetView の下記のプロパティのいずれかが true の場合に発生します。
  - AutoGenerateColumns
  - DataAutoCellTypes
  - DataAutoSizeColumns
  - DataAutoHeadings
