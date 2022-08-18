using FarPoint.Win;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using FarPoint.Win.SuperEdit;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// ComboBoxCellType のコンバーターを提供します。
    /// </summary>
    internal class ComboBoxCellTypeConverter : BaseCellTypeConverter
    {
        /// <summary>
        /// 新しい ComboBoxCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public ComboBoxCellTypeConverter(ComboBoxCellType cellType) : base(cellType)
        {

        }

        /// <summary>
        /// セルタイプのプロパティをシリアライズします。
        /// ImageList, ListControl はシリアライズを行いません。
        /// </summary>
        /// <param name="propsObj">プロパティオブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        // NOTE: 下記はシリアライズ化対象外。
        //         ImageList: 複数のオブジェクトが対象になる、デザイナで指定不可なため。
        //         ListControl: 任意の ListBox が対象になる、デザイナで指定不可なため。
        protected override void SerializeProp(JObject propsObj, string[] includeProps)
        {
            base.SerializeProp(propsObj, includeProps);

            var c = CellType as ComboBoxCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.AcceptsArrowKeys))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.AcceptsArrowKeys), c.AcceptsArrowKeys));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.AllowArrowKeysMoveActiveCell))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.AllowArrowKeysMoveActiveCell), c.AllowArrowKeysMoveActiveCell));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.AllowEditorVerticalAlign))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.AllowEditorVerticalAlign), c.AllowEditorVerticalAlign));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.AutoCompleteCustomSource))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.AutoCompleteCustomSource), c.AutoCompleteCustomSource));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.AutoCompleteMode))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.AutoCompleteMode), c.AutoCompleteMode));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.AutoCompleteSource))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.AutoCompleteSource), c.AutoCompleteSource));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.AutoFillAutoCompleteCustomSource))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.AutoFillAutoCompleteCustomSource), c.AutoFillAutoCompleteCustomSource));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.AutoSearch))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.AutoSearch), c.AutoSearch));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.BackgroundImage))))
            {
                ImageConverter imgConverter = new ImageConverter();
                byte[] imageBytes = (byte[])imgConverter.ConvertTo(c.BackgroundImage, typeof(byte[]));
                if (imageBytes.Length == 0)
                {
                    propsObj.Add(new JProperty(nameof(ComboBoxCellType.BackgroundImage), null));
                }
                else
                {
                    propsObj.Add(new JProperty(nameof(ComboBoxCellType.BackgroundImage), System.Convert.ToBase64String(imageBytes)));
                }
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.ButtonAlign))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.ButtonAlign), c.ButtonAlign));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.CharacterCasing))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.CharacterCasing), c.CharacterCasing));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.CharacterSet))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.CharacterSet), c.CharacterSet));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.DoubleClickTextToDropDown))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.DoubleClickTextToDropDown), c.DoubleClickTextToDropDown));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.DropDownOptions))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.DropDownOptions), c.DropDownOptions));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.DropDownWhenStartEditing))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.DropDownWhenStartEditing), c.DropDownWhenStartEditing));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.Editable))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.Editable), c.Editable));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.EditorValue))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.EditorValue), c.EditorValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.ItemData))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.ItemData), c.ItemData));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.Items))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.Items), c.Items));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.ListAlignment))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.ListAlignment), c.ListAlignment));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.ListOffset))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.ListOffset), c.ListOffset));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.ListWidth))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.ListWidth), c.ListWidth));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.MaxDrop))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.MaxDrop), c.MaxDrop));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.MaxLength))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.MaxLength), c.MaxLength));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ComboBoxCellType.StopEditingAfterDropDownItemSelected))))
            {
                propsObj.Add(new JProperty(nameof(ComboBoxCellType.StopEditingAfterDropDownItemSelected), c.StopEditingAfterDropDownItemSelected));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// ImageList, ListControl はデシリアライズを行いません。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         ImageList: 複数のオブジェクトが対象になる、デザイナで指定不可なため。
        //         ListControl: 任意の ListBox が対象になる、デザイナで指定不可なため。
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as ComboBoxCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(ComboBoxCellType.AcceptsArrowKeys), true) == 0)
            {
                c.AcceptsArrowKeys = prop.Value.ToObject<AcceptsArrowKeys>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.AllowArrowKeysMoveActiveCell), true) == 0)
            {
                c.AllowArrowKeysMoveActiveCell = prop.Value.ToObject<FlagArrowKeys>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.AllowEditorVerticalAlign), true) == 0)
            {
                c.AllowEditorVerticalAlign = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.AutoCompleteCustomSource), true) == 0)
            {
                c.AutoCompleteCustomSource = prop.Value.ToObject<AutoCompleteStringCollection>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.AutoCompleteMode), true) == 0)
            {
                c.AutoCompleteMode = prop.Value.ToObject<AutoCompleteMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.AutoCompleteSource), true) == 0)
            {
                c.AutoCompleteSource = prop.Value.ToObject<AutoCompleteSource>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.AutoFillAutoCompleteCustomSource), true) == 0)
            {
                c.AutoFillAutoCompleteCustomSource = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.AutoSearch), true) == 0)
            {
                c.AutoSearch = prop.Value.ToObject<AutoSearch>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.BackgroundImage), true) == 0)
            {
                if (prop.Value.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(prop.Value.ToObject<string>());
                    ImageConverter imgConverter = new ImageConverter();
                    var imageObj = (Picture)imgConverter.ConvertFrom(imageValue);
                    c.BackgroundImage = imageObj;
                }
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.ButtonAlign), true) == 0)
            {
                c.ButtonAlign = prop.Value.ToObject<ButtonAlign>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.CharacterCasing), true) == 0)
            {
                c.CharacterCasing = prop.Value.ToObject<CharacterCasing>();
                return;
            }

            if (string.Compare(prop.Key, nameof(ComboBoxCellType.CharacterSet), true) == 0)
            {
                c.CharacterSet = prop.Value.ToObject<ComboCharacterSet>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.DoubleClickTextToDropDown), true) == 0)
            {
                c.DoubleClickTextToDropDown = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.DropDownOptions), true) == 0)
            {
                c.DropDownOptions = prop.Value.ToObject<DropDownOptions>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.DropDownWhenStartEditing), true) == 0)
            {
                c.DropDownWhenStartEditing = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.Editable), true) == 0)
            {
                c.Editable = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.EditorValue), true) == 0)
            {
                c.EditorValue = prop.Value.ToObject<EditorValue>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.ItemData), true) == 0)
            {
                c.ItemData = prop.Value.ToObject<string[]>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.Items), true) == 0)
            {
                c.Items = prop.Value.ToObject<string[]>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.ListAlignment), true) == 0)
            {
                c.ListAlignment = prop.Value.ToObject<ListAlignment>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.ListOffset), true) == 0)
            {
                c.ListOffset = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.ListWidth), true) == 0)
            {
                c.ListWidth = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.MaxDrop), true) == 0)
            {
                c.MaxDrop = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.MaxLength), true) == 0)
            {
                c.MaxLength = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ComboBoxCellType.StopEditingAfterDropDownItemSelected), true) == 0)
            {
                c.StopEditingAfterDropDownItemSelected = prop.Value.ToObject<bool>();
                return;
            }
        }
    }
}
