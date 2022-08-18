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
    /// EditBaseCellType のコンバーターを提供します。
    /// </summary>
    internal class EditBaseCellTypeConverter : BaseCellTypeConverter
    {
        /// <summary>
        /// 新しい EditBaseCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public EditBaseCellTypeConverter(EditBaseCellType cellType) : base(cellType)
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

            var c = CellType as EditBaseCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.AcceptsArrowKeys))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.AcceptsArrowKeys), c.AcceptsArrowKeys));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.AllowArrowKeysMoveActiveCell))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.AllowArrowKeysMoveActiveCell), c.AllowArrowKeysMoveActiveCell));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.AllowEditorVerticalAlign))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.AllowEditorVerticalAlign), c.AllowEditorVerticalAlign));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.AutoCompleteCustomSource))))   // TODO
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.AutoCompleteCustomSource), c.AutoCompleteCustomSource));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.AutoCompleteMode))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.AutoCompleteMode), c.AutoCompleteMode));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.AutoCompleteSource))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.AutoCompleteSource), c.AutoCompleteSource));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.AutoFillAutoCompleteCustomSource))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.AutoFillAutoCompleteCustomSource), c.AutoFillAutoCompleteCustomSource));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.BackgroundImage))))
            {
                ImageConverter imgConverter = new ImageConverter();
                byte[] imageBytes = (byte[])imgConverter.ConvertTo(c.BackgroundImage, typeof(byte[]));
                if (imageBytes.Length == 0)
                {
                    propsObj.Add(new JProperty(nameof(EditBaseCellType.BackgroundImage), null));
                }
                else
                {
                    propsObj.Add(new JProperty(nameof(EditBaseCellType.BackgroundImage), System.Convert.ToBase64String(imageBytes)));
                }
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.ButtonAlign))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.ButtonAlign), c.ButtonAlign));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.DropDownButton))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.DropDownButton), c.DropDownButton));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.EnableSubEditor))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.EnableSubEditor), c.EnableSubEditor));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.FocusPosition))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.FocusPosition), c.FocusPosition));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.NullDisplay))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.NullDisplay), c.NullDisplay));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.ReadOnly))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.ReadOnly), c.ReadOnly));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.ShrinkToFit))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.ShrinkToFit), c.ShrinkToFit));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.Static))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.Static), c.Static));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.StringTrim))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.StringTrim), c.StringTrim));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.TextOrientation))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.TextOrientation), c.TextOrientation));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.TextRotationAngle))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.TextRotationAngle), c.TextRotationAngle));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(EditBaseCellType.WordWrap))))
            {
                propsObj.Add(new JProperty(nameof(EditBaseCellType.WordWrap), c.WordWrap));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as EditBaseCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(EditBaseCellType.AcceptsArrowKeys), true) == 0)
            {
                c.AcceptsArrowKeys = prop.Value.ToObject<AcceptsArrowKeys>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.AllowArrowKeysMoveActiveCell), true) == 0)
            {
                c.AllowArrowKeysMoveActiveCell = prop.Value.ToObject<FlagArrowKeys>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.AllowEditorVerticalAlign), true) == 0)
            {
                c.AllowEditorVerticalAlign = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.AutoCompleteCustomSource), true) == 0)
            {
                c.AutoCompleteCustomSource = prop.Value.ToObject<AutoCompleteStringCollection>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.AutoCompleteMode), true) == 0)
            {
                c.AutoCompleteMode = prop.Value.ToObject<AutoCompleteMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.AutoCompleteSource), true) == 0)
            {
                c.AutoCompleteSource = prop.Value.ToObject<AutoCompleteSource>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.AutoFillAutoCompleteCustomSource), true) == 0)
            {
                c.AutoFillAutoCompleteCustomSource = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.BackgroundImage), true) == 0)
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
            if (string.Compare(prop.Key, nameof(EditBaseCellType.ButtonAlign), true) == 0)
            {
                c.ButtonAlign = prop.Value.ToObject<ButtonAlign>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.DropDownButton), true) == 0)
            {
                c.DropDownButton = prop.Value.ToObject<bool>();
                return;
            }

            if (string.Compare(prop.Key, nameof(EditBaseCellType.EnableSubEditor), true) == 0)
            {
                c.EnableSubEditor = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.FocusPosition), true) == 0)
            {
                c.FocusPosition = prop.Value.ToObject<EditorFocusCursorPosition>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.NullDisplay), true) == 0)
            {
                c.NullDisplay = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.ReadOnly), true) == 0)
            {
                c.ReadOnly = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.ShrinkToFit), true) == 0)
            {
                c.ShrinkToFit = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.Static), true) == 0)
            {
                c.Static = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.StringTrim), true) == 0)
            {
                c.StringTrim = prop.Value.ToObject<StringTrimming>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.TextOrientation), true) == 0)
            {
                c.TextOrientation = prop.Value.ToObject<TextOrientation>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.TextRotationAngle), true) == 0)
            {
                c.TextRotationAngle = prop.Value.ToObject<double>();
                return;
            }
            if (string.Compare(prop.Key, nameof(EditBaseCellType.WordWrap), true) == 0)
            {
                c.WordWrap = prop.Value.ToObject<bool>();
                return;
            }
        }
    }
}
