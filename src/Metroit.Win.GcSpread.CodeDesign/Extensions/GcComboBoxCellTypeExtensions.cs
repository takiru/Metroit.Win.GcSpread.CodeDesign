using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters;

namespace Metroit.Win.GcSpread.CodeDesign.Extensions
{
    /// <summary>
    /// GcComboBoxCellType のJSONシリアライズ、デシリアライズを提供します。
    /// </summary>
    public static class GcComboBoxCellTypeExtensions
    {
        /// <summary>
        /// GcComboBoxCellType をJSONシリアライズします。
        /// DataSource, ImageList, Items, ListColumns[].DefaultSubItem, ListColumns[].SortComparer, SubEditor, TouchToolBar はシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">GcComboBoxCellType オブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        /// <returns>JSON文字列。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         DataSource: 様々なオブジェクトが指定されるため。
        //         ImageList: 複数のオブジェクトが対象になる、デザイナで指定不可なため。
        //         Items: DataSource 指定もしくは直接指定があるため。
        //         ListColumns[].DefaultSubItem: 既定値設定？のため。
        //         ListColumns[].SortComparer: ソート制御用クラスのため。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        public static string SerializeJson(this GcComboBoxCellType cellType, string[] includeProps = null)
        {
            var converter = new GcComboBoxCellTypeConverter(cellType);
            return converter.SerializeJson(includeProps);
        }

        /// <summary>
        /// GcComboBoxCellType をJSONデシリアライズします。
        /// DataSource, ImageList, Items, ListColumns[].DefaultSubItem, ListColumns[].SortComparer, SubEditor, TouchToolBar はデシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">GcComboBoxCellType オブジェクト。</param>
        /// <param name="cellTypeProps">CellType のプロパティのJSON文字列。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         DataSource: 様々なオブジェクトが指定されるため。
        //         ImageList: 複数のオブジェクトが対象になる、デザイナで指定不可なため。
        //         Items: DataSource 指定もしくは直接指定があるため。
        //         ListColumns[].DefaultSubItem: 既定値設定？のため。
        //         ListColumns[].SortComparer: ソート制御用クラスのため。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        public static void DeserializeJson(this GcComboBoxCellType cellType, string cellTypeProps)
        {
            var converter = new GcComboBoxCellTypeConverter(cellType);
            converter.DeserializeJson(cellTypeProps);
        }
    }
}
