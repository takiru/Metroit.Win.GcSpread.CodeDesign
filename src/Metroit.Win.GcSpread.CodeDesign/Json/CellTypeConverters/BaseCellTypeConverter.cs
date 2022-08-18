using FarPoint.Win.Spread.CellType;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// BaseCellType のコンバーターを提供します。
    /// </summary>
    internal class BaseCellTypeConverter : CellTypeConverterBase
    {
        /// <summary>
        /// 新しい BaseCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public BaseCellTypeConverter(BaseCellType cellType) : base(cellType)
        {

        }

        /// <summary>
        /// セルタイプのプロパティをシリアライズします。
        /// SubEditor はシリアライズを行いません。
        /// </summary>
        /// <param name="propsObj">プロパティオブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        // NOTE: 下記はシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        protected override void SerializeProp(JObject propsObj, string[] includeProps)
        {
            return;
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// SubEditor はデシリアライズを行いません。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            return;
        }
    }
}
