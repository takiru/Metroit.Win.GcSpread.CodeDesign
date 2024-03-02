using FarPoint.Win.Spread.CellType;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// RegularExpressionCellType のコンバーターを提供します。
    /// </summary>
    internal class RegularExpressionCellTypeConverter : EditBaseCellTypeConverter
    {
        /// <summary>
        /// 新しい RegularExpressionCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public RegularExpressionCellTypeConverter(RegularExpressionCellType cellType) : base(cellType)
        {

        }

        /// <summary>
        /// セルタイプのプロパティをシリアライズします。
        /// </summary>
        /// <param name="propsObj">プロパティオブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        protected override void SerializeProp(JObject propsObj, string[] includeProps)
        {
            base.SerializeProp(propsObj, includeProps);

            var c = CellType as RegularExpressionCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x =>x.Contains(nameof(RegularExpressionCellType.RegularExpression))))
            {
                propsObj.Add(new JProperty(nameof(RegularExpressionCellType.RegularExpression), c.RegularExpression));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as RegularExpressionCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(RegularExpressionCellType.RegularExpression), true) == 0)
            {
                c.RegularExpression = prop.Value.ToObject<string>();
                return;
            }
        }
    }
}
