using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign;
using Metroit.Win.GcSpread.CodeDesign.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Test
{
    public partial class SampleForm : Form
    {
        public SampleForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
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
                new ListItem() { Column1 = DateTime.Today, Column2=123, Column3="Row1", Column4="item1", Column5="aaa", Column6="xxx", Column7=1 },
                new ListItem() { Column1 = DateTime.Today, Column2=123, Column3="Row2", Column4="item2", Column5="bbb", Column6="yyy", Column7=0 },
                new ListItem() { Column1 = DateTime.Today, Column2=123, Column3="Row3", Column4="item3", Column5="ccc", Column6="zzz", Column7=1 }
            };

            // JSONからレイアウトオブジェクトを取得
            var json = SampleRex.SampleJson;
            var root = JsonConvert.DeserializeSheetView(json);

            // レイアウトオブジェクトからレイアウトをバインドする
            fpSpread1.ActiveSheet.BindJsonLayout(root,
                null,
                (sheet, root2) =>
                {
                    // ここで、ユーザー単位の列位置、タイトルなどで書き換えることができる。
                    // Columns[].Options プロパティを Tag に入れるなどもここで行う。
                    // 列位置は下記コードによって行い、列移動結果を返却する必要がある。
                    // 列移動がない場合は戻り値は null でよい。

                    var result = new List<ColumnMoveResult>();

                    // 1列目1行目ヘッダーのタイトルを変更する。
                    sheet.ColumnHeader.Cells[0, 0].Value = "値変更";

                    // Options を Tag に入れる
                    foreach (var column in root2.Columns.Select((Item, Index) => new { Item, Index }))
                    {
                        sheet.Columns[column.Index].Tag = column.Item.Options;
                    }

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
                        var cmrs = sheet.MoveColumnKeepDataField(columnMove.beforeIndex, columnMove.afterIndex, columnMove.columnCount);

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
            var comboBoxColumn = fpSpread1.ActiveSheet.Columns.Cast<Column>().Where(x => x.DataField == "Column4").FirstOrDefault();
            var comboBoxCellType = (ComboBoxCellType)comboBoxColumn.CellType;
            comboBoxCellType.Items = new string[] { "item1", "item2", "item3" };
            comboBoxCellType.ItemData = new string[] { "value1", "value2", "value3" };

            // GcComboBox のアイテム設定
            var cellType = fpSpread1.ActiveSheet.Columns[10].CellType as GcComboBoxCellType;

            cellType.DropDownStyle = ComboBoxStyle.DropDownList;
            cellType.EditorValue = GcComboBoxEditorValue.Value; // 実際の値はValueとする

            // 列や行の境界線が邪魔なので消す
            cellType.ListGridLines.HorizontalLines.Style = LineStyle.None;
            cellType.ListGridLines.VerticalLines.Style = LineStyle.None;

            cellType.ListHeaderPane.Visible = false;    // ヘッダー列名を表示しない
            cellType.DropDown.AutoWidth = true; // 列非表示に伴い、幅調整されるようにする
            cellType.DropDown.AllowResize = false;  // ユーザーによってドロップダウンのサイズを変更することを許可しない
            cellType.UseCompatibleDrawing = true;   // ドロップダウンを開いているときに入力エリア部分が灰色にならないようにする

            cellType.AutoGenerateColumns = true;
            cellType.DataSource = comboItems;

            // 内部の値と画面に表示される値を設定
            cellType.ValueSubItemIndex = cellType.ListColumns.IndexOf(cellType.ListColumns.Cast<ListColumnInfo>().Where(x => x.DataPropertyName == "Value").First());
            cellType.TextSubItemIndex = cellType.ListColumns.IndexOf(cellType.ListColumns.Cast<ListColumnInfo>().Where(x => x.DataPropertyName == "Display").First());

            cellType.ListDefaultColumn.Visible = false; // 何もリスト表示しない状態をデフォルトにする
            cellType.ListColumns.Cast<ListColumnInfo>().Where(x => x.DataPropertyName == "Display").First().Visible = true;  // 表示したい列だけ表示する



            fpSpread1.ActiveSheet.DataSource = dataSource;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var root = fpSpread1.ActiveSheet.CreateSpecialColumnDefinitions();

            var json = JsonConvert.SerializeSpecialColumn(root);
            var deserialize = JsonConvert.DeserializeSpecialColumn(json);
            MessageBox.Show(json);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var root = fpSpread1.ActiveSheet.CreateSpecialColumnHeaderDefinitions();
            var json = JsonConvert.SerializeSpecialColumnHeader(root);
            var deserialize = JsonConvert.DeserializeSpecialColumnHeader(json);
            MessageBox.Show(json);
        }

        private void button4_Click(object sender, EventArgs e)
        {
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
            var json = JsonConvert.SerializeSheetView(root);
            MessageBox.Show(json);
        }
    }

    public class ListItem
    {
        public DateTime Column1 { get; set; }

        public long Column2 { get; set; }

        public string Column3 { get; set; }

        public string Column4 { get; set; }

        public string Column5 { get; set; }

        public string Column6 { get; set; }

        public int Column7 { get; set; }
    }

    public class ComboItem
    {
        public string Display { get; set; }

        public string Value { get; set; }
    }
}
