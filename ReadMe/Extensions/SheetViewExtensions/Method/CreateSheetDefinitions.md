# SheetViewExtensions.CreateSheetDefinitions メソッド
# 定義
名前空間: Metroit.Win.GcSpread.CodeDesign

シート情報からJSONデータオブジェクトを生成します。

# CreateSheetDefinitions(Func<FarPoint.Win.Spread.Column, Dictionary<string, object>>)
シート情報からJSONデータオブジェクトを生成します。
```cs
public static SheetViewDefinitions CreateSheetDefinitions(Func<FarPoint.Win.Spread.Column, Dictionary<string, object>> columnOptions = null)
```

### パラメーター
columnOptions  
列のオプション情報。

### 戻り値
SheetViewDefinitions  
SheetViewDefinitions オブジェクト。

### columnOptions
意図的に jsonデータを超越して制御を行いたい場合に指定します。  
アプリケーション固有の値などをオプションとして返却することができます。  

```json
var root = fpSpread1.ActiveSheet.CreateSheetDefinitions((column) =>
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
```