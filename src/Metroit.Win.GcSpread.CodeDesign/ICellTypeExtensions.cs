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



            // InputManセルタイプ
            if (cellType is GcTextBoxCellType)
            {
                ((GcTextBoxCellType)cellType).DeserializeJson(cellTypeProps);
            }
            if (cellType is GcDateTimeCellType)
            {
                ((GcDateTimeCellType)cellType).DeserializeJson(cellTypeProps);
            }
            if (cellType is GcNumberCellType)
            {
                ((GcNumberCellType)cellType).DeserializeJson(cellTypeProps);
            }
            if (cellType is GcComboBoxCellType)
            {
                ((GcComboBoxCellType)cellType).DeserializeJson(cellTypeProps);
            }
        }

        /// <summary>
        /// セルタイプごとにオブジェクトをデシリアライズします。
        /// </summary>
        /// <param name="cellType"></param>
        public static JObject SerializeJson(this ICellType cellType)
        {
            // 通常セルタイプ



            // InputManセルタイプ
            if (cellType is GcTextBoxCellType)
            {
                return JObject.Parse(((GcTextBoxCellType)cellType).SerializeJson());
            }
            if (cellType is GcDateTimeCellType)
            {
                return JObject.Parse(((GcDateTimeCellType)cellType).SerializeJson());
            }
            if (cellType is GcNumberCellType)
            {
                return JObject.Parse(((GcNumberCellType)cellType).SerializeJson());
            }
            if (cellType is GcComboBoxCellType)
            {
                return JObject.Parse(((GcComboBoxCellType)cellType).SerializeJson());
            }

            return null;
        }
    }
}
