using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.Converters;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// GcTextBoxCellType のコンバーターを提供します。
    /// </summary>
    internal class GcTextBoxCellTypeConverter : InputManCellTypeBaseConverter
    {
        /// <summary>
        /// 新しい GcTextBoxCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public GcTextBoxCellTypeConverter(GcTextBoxCellType cellType) : base(cellType)
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

            var c = CellType as GcTextBoxCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.AcceptsCrLf))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.AcceptsCrLf), c.AcceptsCrLf));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.AcceptsTabChar))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.AcceptsTabChar), c.AcceptsTabChar));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.AllowSpace))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.AllowSpace), c.AllowSpace));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.AlternateText))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.AlternateText), TextBoxAlternateTextInfoConverter.Serialize(c.AlternateText)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.AutoComplete))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.AutoComplete), AutoCompleteInfoConverter.Serialize(c.AutoComplete)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.AutoCompleteCustomSource))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.AutoCompleteCustomSource), c.AutoCompleteCustomSource));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.AutoCompleteMode))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.AutoCompleteMode), c.AutoCompleteMode));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.AutoCompleteSource))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.AutoCompleteSource), c.AutoCompleteSource));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.AutoConvert))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.AutoConvert), c.AutoConvert));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.DisplayAlignment))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.DisplayAlignment), c.DisplayAlignment));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.DropDown))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.DropDown), DropDownInfoConverter.Serialize(c.DropDown)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.DropDownEditor))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.DropDownEditor), DropDownEditorInfoConverter.Serialize(c.DropDownEditor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.Ellipsis))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.Ellipsis), c.Ellipsis));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.EllipsisString))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.EllipsisString), c.EllipsisString));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.FocusPosition))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.FocusPosition), c.FocusPosition));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.FormatString))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.FormatString), c.FormatString));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.GridLine))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.GridLine), LineConverter.Serialize(c.GridLine)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.LineSpace))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.LineSpace), c.LineSpace));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.MaxLength))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.MaxLength), c.MaxLength));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.MaxLengthUnit))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.MaxLengthUnit), c.MaxLengthUnit));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.MaxLineCount))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.MaxLineCount), c.MaxLineCount));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.Multiline))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.Multiline), c.Multiline));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.PasswordChar))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.PasswordChar), c.PasswordChar.ToString()));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.RecommendedValue))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.RecommendedValue), c.RecommendedValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.ScrollBarMode))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.ScrollBarMode), c.ScrollBarMode));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.ScrollBars))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.ScrollBars), c.ScrollBars));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.ShowRecommendedValue))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.ShowRecommendedValue), c.ShowRecommendedValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.SideButtons))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.SideButtons), SideButtonCollectionInfoConverter.Serialize(c.SideButtons)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.UseSystemPasswordChar))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.UseSystemPasswordChar), c.UseSystemPasswordChar));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcTextBoxCellType.WrapMode))))
            {
                propsObj.Add(new JProperty(nameof(GcTextBoxCellType.WrapMode), c.WrapMode));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as GcTextBoxCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AcceptsCrLf), true) == 0)
            {
                c.AcceptsCrLf = prop.Value.ToObject<CrLfMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AcceptsTabChar), true) == 0)
            {
                c.AcceptsTabChar = prop.Value.ToObject<TabCharMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AllowSpace), true) == 0)
            {
                c.AllowSpace = prop.Value.ToObject<AllowSpace>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AlternateText), true) == 0)
            {
                TextBoxAlternateTextInfoConverter.Deserialize(c.AlternateText, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AutoComplete), true) == 0)
            {
                AutoCompleteInfoConverter.Deserialize(c.AutoComplete, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AutoCompleteCustomSource), true) == 0)
            {
                c.AutoCompleteCustomSource = prop.Value.ToObject<AutoCompleteStringCollection>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AutoCompleteMode), true) == 0)
            {
                c.AutoCompleteMode = prop.Value.ToObject<AutoCompleteMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AutoCompleteSource), true) == 0)
            {
                c.AutoCompleteSource = prop.Value.ToObject<AutoCompleteSource>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AutoConvert), true) == 0)
            {
                c.AutoConvert = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.DisplayAlignment), true) == 0)
            {
                c.DisplayAlignment = prop.Value.ToObject<DisplayAlignment>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.DropDown), true) == 0)
            {
                DropDownInfoConverter.Deserialize(c.DropDown, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.DropDownEditor), true) == 0)
            {
                DropDownEditorInfoConverter.Deserialize(c.DropDownEditor, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.Ellipsis), true) == 0)
            {
                c.Ellipsis = prop.Value.ToObject<EllipsisMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.EllipsisString), true) == 0)
            {
                c.EllipsisString = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.FocusPosition), true) == 0)
            {
                c.FocusPosition = prop.Value.ToObject<EditorBaseFocusCursorPosition>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.FormatString), true) == 0)
            {
                c.FormatString = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.GridLine), true) == 0)
            {
                c.GridLine = LineConverter.Deserialize(prop.Value, c.GridLine);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.LineSpace), true) == 0)
            {
                c.LineSpace = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.MaxLength), true) == 0)
            {
                c.MaxLength = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.MaxLengthUnit), true) == 0)
            {
                c.MaxLengthUnit = prop.Value.ToObject<LengthUnit>();
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.MaxLineCount), true) == 0)
            {
                c.MaxLineCount = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.Multiline), true) == 0)
            {
                c.Multiline = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.PasswordChar), true) == 0)
            {
                c.PasswordChar = prop.Value.ToObject<char>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.RecommendedValue), true) == 0)
            {
                c.RecommendedValue = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ScrollBarMode), true) == 0)
            {
                c.ScrollBarMode = prop.Value.ToObject<ScrollBarMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ScrollBars), true) == 0)
            {
                c.ScrollBars = prop.Value.ToObject<ScrollBars>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ShowRecommendedValue), true) == 0)
            {
                c.ShowRecommendedValue = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.SideButtons), true) == 0)
            {
                c.SideButtons = SideButtonCollectionInfoConverter.Deserialize((JArray)prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.UseSystemPasswordChar), true) == 0)
            {
                c.UseSystemPasswordChar = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcTextBoxCellType.WrapMode), true) == 0)
            {
                c.WrapMode = prop.Value.ToObject<WrapMode>();
                return;
            }
        }
    }
}
