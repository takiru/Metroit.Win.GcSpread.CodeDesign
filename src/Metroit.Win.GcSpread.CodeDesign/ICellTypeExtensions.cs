using FarPoint.Win.Spread.CellType;
using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign
{
    /// <summary>
    /// セルタイプの拡張メソッドを提供します。
    /// </summary>
    internal static class ICellTypeExtensions
    {
        /// <summary>
        /// セルタイプごとにオブジェクトをデシリアライズします。
        /// </summary>
        /// <param name="cellType"></param>
        /// <param name="cellTypeProps"></param>
        public static void DeserializeJson(this ICellType cellType, string cellTypeProps)
        {
            // 通常セルタイプ
            if (cellType is TextCellType)
            {
                ((TextCellType)cellType).DeserializeJson(cellTypeProps);
                return;
            }
            if (cellType is DateTimeCellType)
            {
                ((DateTimeCellType)cellType).DeserializeJson(cellTypeProps);
                return;
            }
            if (cellType is NumberCellType)
            {
                ((NumberCellType)cellType).DeserializeJson(cellTypeProps);
                return;
            }
            if (cellType is ComboBoxCellType)
            {
                ((ComboBoxCellType)cellType).DeserializeJson(cellTypeProps);
                return;
            }
            if (cellType is CheckBoxCellType)
            {
                ((CheckBoxCellType)cellType).DeserializeJson(cellTypeProps);
                return;
            }
            if (cellType is ButtonCellType)
            {
                ((ButtonCellType)cellType).DeserializeJson(cellTypeProps);
                return;
            }

            // InputManセルタイプ
            if (cellType is GcTextBoxCellType)
            {
                ((GcTextBoxCellType)cellType).DeserializeJson(cellTypeProps);
                return;
            }
            if (cellType is GcDateTimeCellType)
            {
                ((GcDateTimeCellType)cellType).DeserializeJson(cellTypeProps);
                return;
            }
            if (cellType is GcNumberCellType)
            {
                ((GcNumberCellType)cellType).DeserializeJson(cellTypeProps);
                return;
            }
            if (cellType is GcComboBoxCellType)
            {
                ((GcComboBoxCellType)cellType).DeserializeJson(cellTypeProps);
                return;
            }
        }

        /// <summary>
        /// セルタイプごとにオブジェクトをデシリアライズします。
        /// </summary>
        /// <param name="cellType"></param>
        public static JObject SerializeJson(this ICellType cellType, string[] includeProps = null)
        {
            // 通常セルタイプ
            if (cellType is TextCellType)
            {
                return JObject.Parse(((TextCellType)cellType).SerializeJson(includeProps));
            }
            if (cellType is DateTimeCellType)
            {
                return JObject.Parse(((DateTimeCellType)cellType).SerializeJson(includeProps));
            }
            if (cellType is NumberCellType)
            {
                return JObject.Parse(((NumberCellType)cellType).SerializeJson(includeProps));
            }
            if (cellType is ComboBoxCellType)
            {
                return JObject.Parse(((ComboBoxCellType)cellType).SerializeJson(includeProps));
            }
            if (cellType is CheckBoxCellType)
            {
                return JObject.Parse(((CheckBoxCellType)cellType).SerializeJson(includeProps));
            }
            if (cellType is ButtonCellType)
            {
                return JObject.Parse(((ButtonCellType)cellType).SerializeJson(includeProps));
            }

            // InputManセルタイプ
            if (cellType is GcTextBoxCellType)
            {
                return JObject.Parse(((GcTextBoxCellType)cellType).SerializeJson(includeProps));
            }
            if (cellType is GcDateTimeCellType)
            {
                return JObject.Parse(((GcDateTimeCellType)cellType).SerializeJson(includeProps));
            }
            if (cellType is GcNumberCellType)
            {
                return JObject.Parse(((GcNumberCellType)cellType).SerializeJson(includeProps));
            }
            if (cellType is GcComboBoxCellType)
            {
                return JObject.Parse(((GcComboBoxCellType)cellType).SerializeJson(includeProps));
            }

            return null;
        }
    }
}
