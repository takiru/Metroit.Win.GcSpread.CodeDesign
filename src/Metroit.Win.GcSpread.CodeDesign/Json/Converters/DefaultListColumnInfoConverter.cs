using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DefaultListColumnInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DefaultListColumnInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="defaultListColumnInfo">DefaultListColumnInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(DefaultListColumnInfo defaultListColumnInfo)
        {
            if (defaultListColumnInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(DefaultListColumnInfo.AutoWidth), defaultListColumnInfo.AutoWidth));
            jObj.Add(new JProperty(nameof(DefaultListColumnInfo.DataDisplayType), defaultListColumnInfo.DataDisplayType));
            jObj.Add(new JProperty(nameof(DefaultListColumnInfo.DefaultSubItem), DefaultSubItemInfoConverter.Serialize(defaultListColumnInfo.DefaultSubItem)));
            jObj.Add(new JProperty(nameof(DefaultListColumnInfo.Header), ListHeaderInfoConverter.Serialize(defaultListColumnInfo.Header)));
            jObj.Add(new JProperty(nameof(DefaultListColumnInfo.SortOrder), defaultListColumnInfo.SortOrder));
            jObj.Add(new JProperty(nameof(DefaultListColumnInfo.Visible), defaultListColumnInfo.Visible));
            jObj.Add(new JProperty(nameof(DefaultListColumnInfo.Width), defaultListColumnInfo.Width));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>DefaultListColumnInfo オブジェクト。</returns>
        public static DefaultListColumnInfo Deserialize(JToken prop)
        {
            return Deserialize(prop, new DefaultListColumnInfo());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする DefaultListColumnInfo オブジェクト。</param>
        /// <returns>DefaultListColumnInfo オブジェクト。</returns>
        public static DefaultListColumnInfo Deserialize(JToken prop, DefaultListColumnInfo source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new DefaultListColumnInfo();
            if (source != null)
            {
                result.AutoWidth = source.AutoWidth;
                result.DataDisplayType = source.DataDisplayType;
                result.DefaultSubItem = source.DefaultSubItem;
                result.Header = source.Header;
                result.SortOrder = source.SortOrder;
                result.Visible = source.Visible;
                result.Width = source.Width;
            }

            var autoWidth = prop.SelectToken(nameof(DefaultListColumnInfo.AutoWidth));
            if (autoWidth != null)
            {
                result.AutoWidth = autoWidth.ToObject<bool>();
            }

            var dataDisplayType = prop.SelectToken(nameof(DefaultListColumnInfo.DataDisplayType));
            if (dataDisplayType != null)
            {
                result.DataDisplayType = dataDisplayType.ToObject<DataDisplayType>();
            }

            var defaultSubItem = prop.SelectToken(nameof(DefaultListColumnInfo.DefaultSubItem));
            if (defaultSubItem != null)
            {
                result.DefaultSubItem = DefaultSubItemInfoConverter.Deserialize(defaultSubItem);
            }

            var header = prop.SelectToken(nameof(DefaultListColumnInfo.Header));
            if (header != null)
            {
                result.Header = ListHeaderInfoConverter.Deserialize(header, result.Header);
            }

            var sortOrder = prop.SelectToken(nameof(DefaultListColumnInfo.SortOrder));
            if (sortOrder != null)
            {
                result.SortOrder = sortOrder.ToObject<SortOrder>();
            }

            var width = prop.SelectToken(nameof(DefaultListColumnInfo.Width));
            if (width != null)
            {
                result.Width = width.ToObject<int>();
            }

            return result;
        }
    }
}
