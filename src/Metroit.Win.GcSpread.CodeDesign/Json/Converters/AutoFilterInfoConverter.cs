using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// AutoFilterInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class AutoFilterInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="autoFilterInfo">AutoFilterInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(AutoFilterInfo autoFilterInfo)
        {
            if (autoFilterInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(AutoFilterInfo.Enabled), autoFilterInfo.Enabled));
            jObj.Add(new JProperty(nameof(AutoFilterInfo.Interval), autoFilterInfo.Interval));
            jObj.Add(new JProperty(nameof(AutoFilterInfo.MatchingMode), autoFilterInfo.MatchingMode));
            jObj.Add(new JProperty(nameof(AutoFilterInfo.MatchingSource), autoFilterInfo.MatchingSource));
            jObj.Add(new JProperty(nameof(AutoFilterInfo.MaxFilteredItems), autoFilterInfo.MaxFilteredItems));
            jObj.Add(new JProperty(nameof(AutoFilterInfo.MinimumPrefixLength), autoFilterInfo.MinimumPrefixLength));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="autoFilterInfo">AutoFilterInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(AutoFilterInfo autoFilterInfo, JToken prop)
        {
            var enabled = prop.SelectToken(nameof(AutoFilterInfo.Enabled));
            if (enabled != null)
            {
                autoFilterInfo.Enabled = enabled.ToObject<bool>();
            }

            var interval = prop.SelectToken(nameof(AutoFilterInfo.Interval));
            if (interval != null)
            {
                autoFilterInfo.Interval = interval.ToObject<int>();
            }

            var matchingMode = prop.SelectToken(nameof(AutoFilterInfo.MatchingMode));
            if (matchingMode != null)
            {
                autoFilterInfo.MatchingMode = matchingMode.ToObject<AutoCompleteMatchingMode>();
            }

            var matchingSource = prop.SelectToken(nameof(AutoFilterInfo.MatchingSource));
            if (matchingSource != null)
            {
                autoFilterInfo.MatchingSource = matchingMode.ToObject<FilterMatchingSource>();
            }

            var maxFilteredItems = prop.SelectToken(nameof(AutoFilterInfo.MaxFilteredItems));
            if (maxFilteredItems != null)
            {
                autoFilterInfo.MaxFilteredItems = maxFilteredItems.ToObject<int>();
            }

            var minimumPrefixLength = prop.SelectToken(nameof(AutoFilterInfo.MinimumPrefixLength));
            if (minimumPrefixLength != null)
            {
                autoFilterInfo.MinimumPrefixLength = minimumPrefixLength.ToObject<int>();
            }
        }
    }
}
