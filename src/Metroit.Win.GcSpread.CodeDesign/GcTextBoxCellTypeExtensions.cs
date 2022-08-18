using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters;

namespace Metroit.Win.GcSpread.CodeDesign
{
    /// <summary>
    /// GcTextBoxCellType のJSONシリアライズ、デシリアライズを提供します。
    /// </summary>
    public static class GcTextBoxCellTypeExtensions
    {
        /// <summary>
        /// GcTextBoxCellType をJSONシリアライズします。
        /// SubEditor, TouchToolBar はシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">GcTextBoxCellType オブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        /// <returns>JSON文字列。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        public static string SerializeJson(this GcTextBoxCellType cellType, string[] includeProps = null)
        {
            var converter = new GcTextBoxCellTypeConverter(cellType);
            return converter.SerializeJson(includeProps);
        }

        /// <summary>
        /// GcTextBoxCellType をJSONデシリアライズします。
        /// SubEditor, TouchToolBar はデシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">GcTextBoxCellType オブジェクト。</param>
        /// <param name="cellTypeProps">CellType のプロパティのJSON文字列。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        public static void DeserializeJson(this GcTextBoxCellType cellType, string cellTypeProps)
        {
            var converter = new GcTextBoxCellTypeConverter(cellType);
            converter.DeserializeJson(cellTypeProps);
        }
    }
}
