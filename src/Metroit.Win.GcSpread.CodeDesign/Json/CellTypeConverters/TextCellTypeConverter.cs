using FarPoint.Win.Spread.CellType;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// TextCellType のコンバーターを提供します。
    /// </summary>
    internal class TextCellTypeConverter : EditBaseCellTypeConverter
    {
        /// <summary>
        /// 新しい TextCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public TextCellTypeConverter(TextCellType cellType) : base(cellType)
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

            var c = CellType as TextCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(TextCellType.CharacterCasing))))
            {
                propsObj.Add(new JProperty(nameof(TextCellType.CharacterCasing), c.CharacterCasing));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(TextCellType.CharacterSet))))
            {
                propsObj.Add(new JProperty(nameof(TextCellType.CharacterSet), c.CharacterSet));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(TextCellType.HotkeyPrefix))))
            {
                propsObj.Add(new JProperty(nameof(TextCellType.HotkeyPrefix), c.HotkeyPrefix));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(TextCellType.MaxLength))))
            {
                propsObj.Add(new JProperty(nameof(TextCellType.MaxLength), c.MaxLength));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(TextCellType.Multiline))))
            {
                propsObj.Add(new JProperty(nameof(TextCellType.Multiline), c.Multiline));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(TextCellType.PasswordChar))))
            {
                propsObj.Add(new JProperty(nameof(TextCellType.PasswordChar), c.PasswordChar.ToString()));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(TextCellType.ScrollBars))))
            {
                propsObj.Add(new JProperty(nameof(TextCellType.ScrollBars), c.ScrollBars));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as TextCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(TextCellType.CharacterCasing), true) == 0)
            {
                c.CharacterCasing = prop.Value.ToObject<CharacterCasing>();
                return;
            }
            if (string.Compare(prop.Key, nameof(TextCellType.CharacterSet), true) == 0)
            {
                c.CharacterSet = prop.Value.ToObject<CharacterSet>();
                return;
            }
            if (string.Compare(prop.Key, nameof(TextCellType.HotkeyPrefix), true) == 0)
            {
                c.HotkeyPrefix = prop.Value.ToObject<HotkeyPrefix>();
                return;
            }
            if (string.Compare(prop.Key, nameof(TextCellType.MaxLength), true) == 0)
            {
                c.MaxLength = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(TextCellType.Multiline), true) == 0)
            {
                c.Multiline = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(TextCellType.PasswordChar), true) == 0)
            {
                c.PasswordChar = prop.Value.ToObject<char>();
                return;
            }
            if (string.Compare(prop.Key, nameof(TextCellType.ScrollBars), true) == 0)
            {
                c.ScrollBars = prop.Value.ToObject<ScrollBars>();
                return;
            }
        }
    }
}
