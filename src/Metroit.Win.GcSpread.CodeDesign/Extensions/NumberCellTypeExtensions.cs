﻿using FarPoint.Win.Spread.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters;

namespace Metroit.Win.GcSpread.CodeDesign.Extensions
{
    /// <summary>
    /// NumberCellType のJSONシリアライズ、デシリアライズを提供します。
    /// </summary>
    public static class NumberCellTypeExtensions
    {
        /// <summary>
        /// NumberCellType をJSONシリアライズします。
        /// SubEditor はシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">NumberCellType オブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        /// <returns>JSON文字列。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        public static string SerializeJson(this NumberCellType cellType, string[] includeProps = null)
        {
            var converter = new NumberCellTypeConverter(cellType);
            return converter.SerializeJson(includeProps);
        }

        /// <summary>
        /// NumberCellType をJSONデシリアライズします。
        /// SubEditor はデシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">NumberCellType オブジェクト。</param>
        /// <param name="cellTypeProps">CellType のプロパティのJSON文字列。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        public static void DeserializeJson(this NumberCellType cellType, string cellTypeProps)
        {
            var converter = new NumberCellTypeConverter(cellType);
            converter.DeserializeJson(cellTypeProps);
        }
    }
}
