using FarPoint.Win;
using FarPoint.Win.Spread.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.Converters;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// CheckBoxCellType のコンバーターを提供します。
    /// </summary>
    internal class CheckBoxCellTypeConverter : BaseCellTypeConverter
    {
        /// <summary>
        /// 新しい CheckBoxCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public CheckBoxCellTypeConverter(CheckBoxCellType cellType) : base(cellType)
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

            var c = CellType as CheckBoxCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(CheckBoxCellType.BackgroundImage))))
            {
                ImageConverter imgConverter = new ImageConverter();
                byte[] imageBytes = (byte[])imgConverter.ConvertTo(c.BackgroundImage, typeof(byte[]));
                if (imageBytes.Length == 0)
                {
                    propsObj.Add(new JProperty(nameof(CheckBoxCellType.BackgroundImage), null));
                }
                else
                {
                    propsObj.Add(new JProperty(nameof(CheckBoxCellType.BackgroundImage), System.Convert.ToBase64String(imageBytes)));
                }
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(CheckBoxCellType.Caption))))
            {
                propsObj.Add(new JProperty(nameof(CheckBoxCellType.Caption), c.Caption));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(CheckBoxCellType.FocusRectangle))))
            {
                propsObj.Add(new JProperty(nameof(CheckBoxCellType.FocusRectangle), c.FocusRectangle));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(CheckBoxCellType.HotkeyPrefix))))
            {
                propsObj.Add(new JProperty(nameof(CheckBoxCellType.HotkeyPrefix), c.HotkeyPrefix));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(CheckBoxCellType.Picture))))
            {
                propsObj.Add(new JProperty(nameof(CheckBoxCellType.Picture), CheckBoxPictureConverter.Serialize(c.Picture)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(CheckBoxCellType.TextAlign))))
            {
                propsObj.Add(new JProperty(nameof(CheckBoxCellType.TextAlign), c.TextAlign));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(CheckBoxCellType.TextFalse))))
            {
                propsObj.Add(new JProperty(nameof(CheckBoxCellType.TextFalse), c.TextFalse));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(CheckBoxCellType.TextIndeterminate))))
            {
                propsObj.Add(new JProperty(nameof(CheckBoxCellType.TextIndeterminate), c.TextIndeterminate));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(CheckBoxCellType.TextTrue))))
            {
                propsObj.Add(new JProperty(nameof(CheckBoxCellType.TextTrue), c.TextTrue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(CheckBoxCellType.ThreeState))))
            {
                propsObj.Add(new JProperty(nameof(CheckBoxCellType.ThreeState), c.ThreeState));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as CheckBoxCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(CheckBoxCellType.BackgroundImage), true) == 0)
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
            if (string.Compare(prop.Key, nameof(CheckBoxCellType.Caption), true) == 0)
            {
                c.Caption = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(CheckBoxCellType.FocusRectangle), true) == 0)
            {
                c.FocusRectangle = prop.Value.ToObject<FocusRectangle>();
                return;
            }
            if (string.Compare(prop.Key, nameof(CheckBoxCellType.HotkeyPrefix), true) == 0)
            {
                c.HotkeyPrefix = prop.Value.ToObject<HotkeyPrefix>();
                return;
            }
            if (string.Compare(prop.Key, nameof(CheckBoxCellType.Picture), true) == 0)
            {
                c.Picture = CheckBoxPictureConverter.Deserialize(prop.Value, c.Picture);
                return;
            }
            if (string.Compare(prop.Key, nameof(CheckBoxCellType.TextAlign), true) == 0)
            {
                c.TextAlign = prop.Value.ToObject<ButtonTextAlign>();
                return;
            }
            if (string.Compare(prop.Key, nameof(CheckBoxCellType.TextFalse), true) == 0)
            {
                c.TextFalse = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(CheckBoxCellType.TextIndeterminate), true) == 0)
            {
                c.TextIndeterminate = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(CheckBoxCellType.TextTrue), true) == 0)
            {
                c.TextTrue = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(CheckBoxCellType.ThreeState), true) == 0)
            {
                c.ThreeState = prop.Value.ToObject<bool>();
                return;
            }
        }
    }
}
