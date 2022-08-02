using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// ShortcutDictionary オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class ShortcutDictionaryConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="shortcutDictionary">ShortcutDictionary オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(ShortcutDictionary shortcutDictionary)
        {
            if (shortcutDictionary == null)
            {
                return null;
            }

            return JObject.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(shortcutDictionary));
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="autoCompleteInfo">ShortcutDictionary オブジェクト。</param>
        /// <param name="props">デシリアライズオブジェクト。</param>
        public static void Deserialize(ShortcutDictionary shortcutDictionary, JToken props)
        {
            foreach (JProperty prop in props)
            {
                Keys targetKey;
                if (!Enum.TryParse(prop.Name, out targetKey))
                {
                    continue;
                }
                shortcutDictionary.Add(targetKey, prop.Value.ToObject<string>());
            }
        }
    }
}
