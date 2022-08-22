using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// ListColumnInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class ListColumnInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// DefaultSubItem, SortComparer はシリアライズを行いません。
        /// </summary>
        /// <param name="listColumnInfo">ListColumnInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         DefaultSubItem: 既定値設定？のため。
        //         SortComparer: ソート制御用クラスのため。
        public static JObject Serialize(ListColumnInfo listColumnInfo)
        {
            if (listColumnInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(ListColumnInfo.AutoWidth), listColumnInfo.AutoWidth));
            jObj.Add(new JProperty(nameof(ListColumnInfo.DataDisplayType), listColumnInfo.DataDisplayType));
            jObj.Add(new JProperty(nameof(ListColumnInfo.DataPropertyName), listColumnInfo.DataPropertyName));
            jObj.Add(new JProperty(nameof(ListColumnInfo.Header), ListHeaderInfoConverter.Serialize(listColumnInfo.Header)));
            jObj.Add(new JProperty(nameof(ListColumnInfo.SortOrder), listColumnInfo.SortOrder));
            jObj.Add(new JProperty(nameof(ListColumnInfo.Visible), listColumnInfo.Visible));
            jObj.Add(new JProperty(nameof(ListColumnInfo.Width), listColumnInfo.Width));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// DefaultSubItem, SortComparer はデシリアライズを行いません。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>ListColumnInfo オブジェクト。</returns>
        // NOTE: 下記はデシリアライズ化対象外。
        //         DefaultSubItem: 既定値設定？のため。
        //         SortComparer: ソート制御用クラスのため。
        public static ListColumnInfo Deserialize(JToken prop)
        {
            return Deserialize(prop, new ListColumnInfo());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// DefaultSubItem, SortComparer はデシリアライズを行いません。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする ListColumnInfo オブジェクト。</param>
        /// <returns>ListColumnInfo オブジェクト。</returns>
        // NOTE: 下記はデシリアライズ化対象外。
        //         DefaultSubItem: 既定値設定？のため。
        //         SortComparer: ソート制御用クラスのため。
        public static ListColumnInfo Deserialize(JToken prop, ListColumnInfo source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new ListColumnInfo();
            if (source != null)
            {
                result.AutoWidth = source.AutoWidth;
                result.DataDisplayType = source.DataDisplayType;
                result.DataPropertyName = source.DataPropertyName;

                // NOTE: 直接代入はできないため、新たなインスタンスとして設定する。
                result.Header = ListHeaderInfoConverter.Deserialize(new JObject(ListHeaderInfoConverter.Serialize(source.Header)));
                result.SortOrder = source.SortOrder;
                result.Visible = source.Visible;
                result.Width = source.Width;
            }

            var autoWidth = prop.SelectToken(nameof(ListColumnInfo.AutoWidth));
            if (autoWidth != null)
            {
                result.AutoWidth = autoWidth.Value<bool>();
            }

            var dataDisplayType = prop.SelectToken(nameof(ListColumnInfo.DataDisplayType));
            if (dataDisplayType != null)
            {
                result.DataDisplayType = dataDisplayType.Value<DataDisplayType>();
            }

            var dataPropertyName = prop.SelectToken(nameof(ListColumnInfo.DataPropertyName));
            if (dataPropertyName != null)
            {
                result.DataPropertyName = dataPropertyName.Value<string>();
            }

            var header = prop.SelectToken(nameof(ListColumnInfo.Header));
            if (header != null)
            {
                result.Header = ListHeaderInfoConverter.Deserialize(header, result.Header);
            }

            var sortOrder = prop.SelectToken(nameof(ListColumnInfo.SortOrder));
            if (sortOrder != null)
            {
                result.SortOrder = sortOrder.Value<SortOrder>();
            }

            var visible = prop.SelectToken(nameof(ListColumnInfo.Visible));
            if (visible != null)
            {
                result.Visible = visible.Value<bool>();
            }

            var width = prop.SelectToken(nameof(ListColumnInfo.Width));
            if (width != null)
            {
                result.Width = width.Value<int>();
            }

            return result;
        }
    }
}
