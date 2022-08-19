using FarPoint.Win.Spread.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters;

namespace Metroit.Win.GcSpread.CodeDesign.Extensions
{
    /// <summary>
    /// ComboBoxCellType のJSONシリアライズ、デシリアライズを提供します。
    /// </summary>
    public static class ComboBoxCellTypeExtensions
    {
        /// <summary>
        /// ComboBoxCellType をJSONシリアライズします。
        /// SubEditor, ImageList, ListControl はシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">ComboBoxCellType オブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        /// <returns>JSON文字列。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         ImageList: 複数のオブジェクトが対象になる、デザイナで指定不可なため。
        //         ListControl: 任意の ListBox が対象になる、デザイナで指定不可なため。
        public static string SerializeJson(this ComboBoxCellType cellType, string[] includeProps = null)
        {
            var converter = new ComboBoxCellTypeConverter(cellType);
            return converter.SerializeJson(includeProps);
        }

        /// <summary>
        /// ComboBoxCellType をJSONデシリアライズします。
        /// SubEditor, ImageList, ListControl はデシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">ComboBoxCellType オブジェクト。</param>
        /// <param name="cellTypeProps">CellType のプロパティのJSON文字列。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         ImageList: 複数のオブジェクトが対象になる、デザイナで指定不可なため。
        //         ListControl: 任意の ListBox が対象になる、デザイナで指定不可なため。
        public static void DeserializeJson(this ComboBoxCellType cellType, string cellTypeProps)
        {
            var converter = new ComboBoxCellTypeConverter(cellType);
            converter.DeserializeJson(cellTypeProps);
        }
    }
}
