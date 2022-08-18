using FarPoint.Win;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// CheckBoxPicture オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class CheckBoxPictureConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="checkBoxPicture">CheckBoxPicture オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(CheckBoxPicture checkBoxPicture)
        {
            if (checkBoxPicture == null)
            {
                return null;
            }

            var jObj = new JObject();

            ImageConverter imgConverter = new ImageConverter();
            byte[] imageBytes;

            imageBytes = (byte[])imgConverter.ConvertTo(checkBoxPicture.False, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.False), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.False), System.Convert.ToBase64String(imageBytes)));
            }

            imageBytes = (byte[])imgConverter.ConvertTo(checkBoxPicture.FalseDisabled, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.FalseDisabled), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.FalseDisabled), System.Convert.ToBase64String(imageBytes)));
            }

            imageBytes = (byte[])imgConverter.ConvertTo(checkBoxPicture.FalsePressed, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.FalsePressed), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.FalsePressed), System.Convert.ToBase64String(imageBytes)));
            }

            imageBytes = (byte[])imgConverter.ConvertTo(checkBoxPicture.Indeterminate, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.Indeterminate), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.Indeterminate), System.Convert.ToBase64String(imageBytes)));
            }

            imageBytes = (byte[])imgConverter.ConvertTo(checkBoxPicture.IndeterminateDisabled, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.IndeterminateDisabled), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.IndeterminateDisabled), System.Convert.ToBase64String(imageBytes)));
            }

            imageBytes = (byte[])imgConverter.ConvertTo(checkBoxPicture.IndeterminatePressed, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.IndeterminatePressed), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.IndeterminatePressed), System.Convert.ToBase64String(imageBytes)));
            }

            imageBytes = (byte[])imgConverter.ConvertTo(checkBoxPicture.True, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.True), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.True), System.Convert.ToBase64String(imageBytes)));
            }

            imageBytes = (byte[])imgConverter.ConvertTo(checkBoxPicture.TrueDisabled, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.TrueDisabled), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.TrueDisabled), System.Convert.ToBase64String(imageBytes)));
            }

            imageBytes = (byte[])imgConverter.ConvertTo(checkBoxPicture.TruePressed, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.TruePressed), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(CheckBoxPicture.TruePressed), System.Convert.ToBase64String(imageBytes)));
            }

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>CheckBoxPicture オブジェクト。</returns>
        public static CheckBoxPicture Deserialize(JToken prop)
        {
            return Deserialize(prop, new CheckBoxPicture());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする CheckBoxPicture オブジェクト。</param>
        /// <returns>CheckBoxPicture オブジェクト。</returns>
        public static CheckBoxPicture Deserialize(JToken prop, CheckBoxPicture source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new CheckBoxPicture();
            if (source != null)
            {
                result.False = source.False;
                result.FalseDisabled = source.FalseDisabled;
                result.FalsePressed = source.FalsePressed;
                result.Indeterminate = source.Indeterminate;
                result.IndeterminateDisabled = source.IndeterminateDisabled;
                result.IndeterminatePressed = source.IndeterminatePressed;
                result.True = source.True;
                result.TrueDisabled = source.TrueDisabled;
                result.TruePressed = source.TruePressed;
            }

            ImageConverter imgConverter = new ImageConverter();

            var falseImage = prop.SelectToken(nameof(CheckBoxPicture.False));
            if (falseImage != null)
            {
                if (falseImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(falseImage.ToObject<string>());
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    result.False = imageObj;
                }
                else
                {
                    result.False = null;
                }
            }

            var falseDisabledImage = prop.SelectToken(nameof(CheckBoxPicture.FalseDisabled));
            if (falseDisabledImage != null)
            {
                if (falseDisabledImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(falseDisabledImage.ToObject<string>());
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    result.FalseDisabled = imageObj;
                }
                else
                {
                    result.FalseDisabled = null;
                }
            }

            var falsePressedImage = prop.SelectToken(nameof(CheckBoxPicture.FalsePressed));
            if (falsePressedImage != null)
            {
                if (falsePressedImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(falsePressedImage.ToObject<string>());
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    result.FalsePressed = imageObj;
                }
                else
                {
                    result.FalsePressed = null;
                }
            }

            var indeterminateImage = prop.SelectToken(nameof(CheckBoxPicture.Indeterminate));
            if (indeterminateImage != null)
            {
                if (indeterminateImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(indeterminateImage.ToObject<string>());
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    result.Indeterminate = imageObj;
                }
                else
                {
                    result.Indeterminate = null;
                }
            }

            var indeterminateDisabledImage = prop.SelectToken(nameof(CheckBoxPicture.IndeterminateDisabled));
            if (indeterminateDisabledImage != null)
            {
                if (indeterminateDisabledImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(indeterminateDisabledImage.ToObject<string>());
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    result.IndeterminateDisabled = imageObj;
                }
                else
                {
                    result.IndeterminateDisabled = null;
                }
            }

            var indeterminatePressedImage = prop.SelectToken(nameof(CheckBoxPicture.IndeterminatePressed));
            if (indeterminatePressedImage != null)
            {
                if (indeterminatePressedImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(indeterminatePressedImage.ToObject<string>());
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    result.IndeterminatePressed = imageObj;
                }
                else
                {
                    result.IndeterminatePressed = null;
                }
            }

            var trueImage = prop.SelectToken(nameof(CheckBoxPicture.True));
            if (trueImage != null)
            {
                if (trueImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(trueImage.ToObject<string>());
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    result.True = imageObj;
                }
                else
                {
                    result.True = null;
                }
            }

            var trueDisabledImage = prop.SelectToken(nameof(CheckBoxPicture.TrueDisabled));
            if (trueDisabledImage != null)
            {
                if (trueDisabledImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(trueDisabledImage.ToObject<string>());
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    result.TrueDisabled = imageObj;
                }
                else
                {
                    result.TrueDisabled = null;
                }
            }

            var truePressedImage = prop.SelectToken(nameof(CheckBoxPicture.TruePressed));
            if (truePressedImage != null)
            {
                if (truePressedImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(truePressedImage.ToObject<string>());
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    result.TruePressed = imageObj;
                }
                else
                {
                    result.TruePressed = null;
                }
            }

            return result;
        }
    }
}
