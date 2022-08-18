using FarPoint.Win.Spread.CellType;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// CellType のコンバーターインターフェースを提供します。
    /// </summary>
    internal abstract class CellTypeConverterBase
    {
        /// <summary>
        /// コンバーターの対象となるセルタイプを取得します。
        /// </summary>
        public ICellType CellType { get; }

        /// <summary>
        /// 新しい CellTypeConverterBase インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public CellTypeConverterBase(ICellType cellType)
        {
            CellType = cellType;
        }

        /// <summary>
        /// セルタイプをシリアライズします。
        /// </summary>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        /// <returns></returns>
        public string SerializeJson(string[] includeProps = null)
        {
            var result = new JObject();
            SerializeProp(result, includeProps);
            return Newtonsoft.Json.JsonConvert.SerializeObject(result);
        }

        /// <summary>
        /// セルタイプをデシリアライズします。
        /// </summary>
        /// <param name="cellTypeProps">セルタイププロパティJSON文字列。</param>
        public void DeserializeJson(string cellTypeProps)
        {
            if (string.IsNullOrEmpty(cellTypeProps))
            {
                return;
            }

            var props = JObject.Parse(cellTypeProps);
            foreach (var prop in props)
            {
                DeserializeProp(prop);
            }
        }

        /// <summary>
        /// セルタイプのプロパティをシリアライズします。
        /// </summary>
        /// <param name="propsObj">プロパティオブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        protected abstract void SerializeProp(JObject propsObj, string[] includeProps);

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected abstract void DeserializeProp(KeyValuePair<string, JToken> prop);
    }
}
