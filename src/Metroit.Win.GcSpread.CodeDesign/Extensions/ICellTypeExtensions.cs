using FarPoint.Win.Spread.CellType;
using GrapeCity.Win.Spread.InputMan.CellType;

namespace Metroit.Win.GcSpread.CodeDesign.Extensions
{
    /// <summary>
    /// セルタイプの拡張メソッドを提供します。
    /// </summary>
    internal static class ICellTypeExtensions
    {
        /// <summary>
        /// セルタイプごとにオブジェクトをデシリアライズします。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        public static string SerializeJson(this ICellType cellType, string[] includeProps = null)
        {
            // 通常セルタイプ
            if (cellType is TextCellType)
            {
                return ((TextCellType)cellType).SerializeJson(includeProps);
            }
            if (cellType is DateTimeCellType)
            {
                return ((DateTimeCellType)cellType).SerializeJson(includeProps);
            }
            if (cellType is NumberCellType)
            {
                return ((NumberCellType)cellType).SerializeJson(includeProps);
            }
            if (cellType is ComboBoxCellType)
            {
                return ((ComboBoxCellType)cellType).SerializeJson(includeProps);
            }
            if (cellType is CheckBoxCellType)
            {
                return ((CheckBoxCellType)cellType).SerializeJson(includeProps);
            }
            if (cellType is ButtonCellType)
            {
                return ((ButtonCellType)cellType).SerializeJson(includeProps);
            }
            if (cellType is RegularExpressionCellType)
            {
                return ((RegularExpressionCellType)cellType).SerializeJson(includeProps);
            }

            // InputManセルタイプ
            if (cellType is GcTextBoxCellType)
            {
                return ((GcTextBoxCellType)cellType).SerializeJson(includeProps);
            }
            if (cellType is GcDateTimeCellType)
            {
                return ((GcDateTimeCellType)cellType).SerializeJson(includeProps);
            }
            if (cellType is GcNumberCellType)
            {
                return((GcNumberCellType)cellType).SerializeJson(includeProps);
            }
            if (cellType is GcComboBoxCellType)
            {
                return ((GcComboBoxCellType)cellType).SerializeJson(includeProps);
            }

            return null;
        }

        /// <summary>
        /// セルタイプごとにオブジェクトをデシリアライズします。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        /// <param name="cellTypeProps">セルタイプのプロパティJSON文字列。</param>
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
            if (cellType is RegularExpressionCellType)
            {
                ((RegularExpressionCellType)cellType).DeserializeJson(cellTypeProps);
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
    }
}
