using FarPoint.Win;
using FarPoint.Win.Spread.CellType;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// ButtonCellType のコンバーターを提供します。
    /// </summary>
    internal class ButtonCellTypeConverter : BaseCellTypeConverter
    {
        /// <summary>
        /// 新しい ButtonCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public ButtonCellTypeConverter(ButtonCellType cellType) : base(cellType)
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

            var c = CellType as ButtonCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.BackgroundStyle))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.BackgroundStyle), c.BackgroundStyle));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.ButtonColor))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.ButtonColor), ColorTranslator.ToHtml(c.ButtonColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.ButtonColor2))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.ButtonColor2), ColorTranslator.ToHtml(c.ButtonColor2)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.DarkColor))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.DarkColor), ColorTranslator.ToHtml(c.DarkColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.GradientMode))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.GradientMode), c.GradientMode));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.HotkeyPrefix))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.HotkeyPrefix), c.HotkeyPrefix));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.LightColor))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.LightColor), ColorTranslator.ToHtml(c.LightColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.Picture))))
            {
                ImageConverter imgConverter = new ImageConverter();
                byte[] imageBytes = (byte[])imgConverter.ConvertTo(c.Picture, typeof(byte[]));
                if (imageBytes.Length == 0)
                {
                    propsObj.Add(new JProperty(nameof(ButtonCellType.Picture), null));
                }
                else
                {
                    propsObj.Add(new JProperty(nameof(ButtonCellType.Picture), System.Convert.ToBase64String(imageBytes)));
                }
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.PictureDown))))
            {
                ImageConverter imgConverter = new ImageConverter();
                byte[] imageBytes = (byte[])imgConverter.ConvertTo(c.PictureDown, typeof(byte[]));
                if (imageBytes.Length == 0)
                {
                    propsObj.Add(new JProperty(nameof(ButtonCellType.PictureDown), null));
                }
                else
                {
                    propsObj.Add(new JProperty(nameof(ButtonCellType.PictureDown), System.Convert.ToBase64String(imageBytes)));
                }
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.ShadowSize))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.ShadowSize), c.ShadowSize));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.Text))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.Text), c.Text));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.TextAlign))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.TextAlign), c.TextAlign));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.TextColor))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.TextColor), ColorTranslator.ToHtml(c.TextColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.TextDown))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.TextDown), c.TextDown));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.TextOrientation))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.TextOrientation), c.TextOrientation));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.TextRotationAngle))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.TextRotationAngle), c.TextRotationAngle));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.TwoState))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.TwoState), c.TwoState));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.UseVisualStyleBackColor))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.UseVisualStyleBackColor), c.UseVisualStyleBackColor));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(ButtonCellType.WordWrap))))
            {
                propsObj.Add(new JProperty(nameof(ButtonCellType.WordWrap), c.WordWrap));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as ButtonCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(ButtonCellType.BackgroundStyle), true) == 0)
            {
                c.BackgroundStyle = prop.Value.ToObject<BackStyle>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.ButtonColor), true) == 0)
            {
                c.ButtonColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.ButtonColor2), true) == 0)
            {
                c.ButtonColor2 = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.DarkColor), true) == 0)
            {
                c.DarkColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.GradientMode), true) == 0)
            {
                c.GradientMode = prop.Value.ToObject<LinearGradientMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.HotkeyPrefix), true) == 0)
            {
                c.HotkeyPrefix = prop.Value.ToObject<HotkeyPrefix>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.LightColor), true) == 0)
            {
                c.LightColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.Picture), true) == 0)
            {
                if (prop.Value.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(prop.Value.ToObject<string>());
                    ImageConverter imgConverter = new ImageConverter();
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    c.Picture = imageObj;
                }
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.PictureDown), true) == 0)
            {
                if (prop.Value.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(prop.Value.ToObject<string>());
                    ImageConverter imgConverter = new ImageConverter();
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    c.PictureDown = imageObj;
                }
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.ShadowSize), true) == 0)
            {
                c.ShadowSize = prop.Value.ToObject<int>();
                return;
            }

            if (string.Compare(prop.Key, nameof(ButtonCellType.Text), true) == 0)
            {
                c.Text = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.TextAlign), true) == 0)
            {
                c.TextAlign = prop.Value.ToObject<ButtonTextAlign>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.TextColor), true) == 0)
            {
                c.TextColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.TextDown), true) == 0)
            {
                c.TextDown = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.TextOrientation), true) == 0)
            {
                c.TextOrientation = prop.Value.ToObject<TextOrientation>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.TextRotationAngle), true) == 0)
            {
                c.TextRotationAngle = prop.Value.ToObject<double>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.TwoState), true) == 0)
            {
                c.TwoState = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.UseVisualStyleBackColor), true) == 0)
            {
                c.UseVisualStyleBackColor = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(ButtonCellType.WordWrap), true) == 0)
            {
                c.WordWrap = prop.Value.ToObject<bool>();
                return;
            }
        }
    }
}
