# Metroit.Win.GcSpread.CodeDesign
GrapeCity SPREAD for Windows Forms のデザインをJSONで行うライブラリです。

# リファレンス
ライブラリの種類・命令は [リファレンス](ReadMe/Reference.md) を読んでください。

# json から SheetView のレイアウトをバインドする例
```cs
// NOTE: この4つは無効でなければならない
fpSpread1.ActiveSheet.AutoGenerateColumns = false;
fpSpread1.ActiveSheet.DataAutoCellTypes = false;
fpSpread1.ActiveSheet.DataAutoSizeColumns = false;
fpSpread1.ActiveSheet.DataAutoHeadings = false;

// コンボボックス用のアイテム
var comboItems = new BindingList<ComboItem>()
{
    new ComboItem() { Display = "item1", Value="value1"},
    new ComboItem() { Display = "item2", Value="value2"},
    new ComboItem() { Display = "item3", Value="value3"}
};

// データソース
var dataSource = new BindingList<ListItem>()
{
    new ListItem() { Column1 = DateTime.Today, Column2=123, Column3="Row1", Column4="item1", Column5="aaa", Column6="xxx" },
    new ListItem() { Column1 = DateTime.Today, Column2=123, Column3="Row2", Column4="item2", Column5="bbb", Column6="yyy" },
    new ListItem() { Column1 = DateTime.Today, Column2=123, Column3="Row3", Column4="item3", Column5="ccc", Column6="zzz" }
};

// JSONからレイアウトオブジェクトを取得(リソースから得た場合)
var json = SampleRex.SampleJson;
var root = JsonConvert.DeserializeObject(json);

// レイアウトオブジェクトからレイアウトをバインドする
fpSpread1.ActiveSheet.BindJsonLayout(root,
    null,
    (sheetView, root2) =>
    {
        // ここで、列位置保持テーブルのユーザー単位の列位置、タイトルなどで書き換える。
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

// ComboBox のアイテムを設定
// 列情報が確定していないため、 BindJsonLayout() より後に行う必要がある
var comboBoxColumn = fpSpread1.ActiveSheet.Columns.Cast<FarPoint.Win.Spread.Column>().Where(x => x.DataField == "Column4").FirstOrDefault();
var comboBoxCellType = (ComboBoxCellType)comboBoxColumn.CellType;
comboBoxCellType.Items = new string[] { "item1", "item2", "item3" };
comboBoxCellType.ItemData = new string[] { "value1", "value2", "value3" };

fpSpread1.ActiveSheet.DataSource = dataSource;
```

# SheetView すべての定義を json に変換する例
```cs
var root = fpSpread1.ActiveSheet.CreateJsonObject((column) =>
{
    // ここでオプション情報の設定などをする
    // オプション情報と仮定
    var options = new List<(string DataField, bool UseFreeWords, bool ExtendWordsEnabled)?>()
        {
            ("Column3", false, false),
            ("Column6", false, false),
            ("Column1", true, true),
            ("Column2", true, true),
            ("Column4", true, true),
            ("Column5", false, false)
        };

    var result = new Dictionary<string, object>();

    var option = options.Where(x => x.Value.DataField == column.DataField).FirstOrDefault();
    if (option == null)
    {
        return result;
    }
    result.Add("UseFreeWords", option.Value.UseFreeWords);
    result.Add("ExtendWordsEnabled", option.Value.ExtendWordsEnabled);

    return result;
});
var json = JsonConvert.SerializeObject(root);
MessageBox.Show(json);
```

# SheetView ヘッダーの特定定義を json に変換する例
```cs
var root = fpSpread1.ActiveSheet.CreateSpecialColumnHeaderJsonObject();
var json = JsonConvert.SerializeSpecialColumnHeaderObject(root);
MessageBox.Show(json);
```

# SheetView 列の特定定義を json に変換する例
```cs
var root = fpSpread1.ActiveSheet.CreateSpecialColumnJsonObject();
var json = JsonConvert.SerializeSpecialColumnObject(root);
MessageBox.Show(json);
```
