using FarPoint.Win.SuperEdit;
using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.Converters;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign
{
    /// <summary>
    /// GcNumberCellType のJSONシリアライズ、デシリアライズを提供します。
    /// </summary>
    public static class GcNumberCellTypeExtensions
    {
        /// <summary>
        /// GcNumberCellType をJSONシリアライズします。
        /// SubEditor, TouchToolBar はシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">GcNumberCellType オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        public static string SerializeJson(this GcNumberCellType cellType)
        {
            var result = new JObject();
            result.Add(new JProperty(nameof(cellType.AcceptsArrowKeys), cellType.AcceptsArrowKeys));
            result.Add(new JProperty(nameof(cellType.AcceptsCrLf), cellType.AcceptsCrLf));
            result.Add(new JProperty(nameof(cellType.AcceptsDecimal), cellType.AcceptsDecimal));
            result.Add(new JProperty(nameof(cellType.AllowDeleteToNull), cellType.AllowDeleteToNull));
            result.Add(new JProperty(nameof(cellType.AlternateText), NumberAlternateTextInfoConverter.Serialize(cellType.AlternateText)));
            result.Add(new JProperty(nameof(cellType.BackgroundImage), PictureConverter.Serialize(cellType.BackgroundImage)));
            result.Add(new JProperty(nameof(cellType.ClipContent), cellType.ClipContent));
            result.Add(new JProperty(nameof(cellType.DisplayFields), NumberDisplayFieldCollectionInfoConverter.Serialize(cellType.DisplayFields)));
            result.Add(new JProperty(nameof(cellType.DropDown), DropDownInfoConverter.Serialize(cellType.DropDown)));
            result.Add(new JProperty(nameof(cellType.DropDownCalculator), DropDownCalculatorInfoConverter.Serialize(cellType.DropDownCalculator)));
            result.Add(new JProperty(nameof(cellType.EditMode), cellType.EditMode));
            result.Add(new JProperty(nameof(cellType.ExcelExportFormat), cellType.ExcelExportFormat));
            result.Add(new JProperty(nameof(cellType.ExitOnLastChar), cellType.ExitOnLastChar));
            result.Add(new JProperty(nameof(cellType.Fields), NumberFieldsInfoConverter.Serialize(cellType.Fields)));
            result.Add(new JProperty(nameof(cellType.FocusPosition), cellType.FocusPosition));
            result.Add(new JProperty(nameof(cellType.MaxMinBehavior), cellType.MaxMinBehavior));
            result.Add(new JProperty(nameof(cellType.MaxValue), cellType.MaxValue));
            result.Add(new JProperty(nameof(cellType.MinValue), cellType.MinValue));
            result.Add(new JProperty(nameof(cellType.NegativeColor), ColorTranslator.ToHtml(cellType.NegativeColor)));
            result.Add(new JProperty(nameof(cellType.PaintByControl), cellType.PaintByControl));
            result.Add(new JProperty(nameof(cellType.PopUpCalculator), PopUpCalculatorInfoConverter.Serialize(cellType.PopUpCalculator)));
            result.Add(new JProperty(nameof(cellType.ReadOnly), cellType.ReadOnly));
            result.Add(new JProperty(nameof(cellType.RecommendedValue), cellType.RecommendedValue));
            result.Add(new JProperty(nameof(cellType.RoundPattern), cellType.RoundPattern));
            result.Add(new JProperty(nameof(cellType.ShortcutKeys), ShortcutDictionaryConverter.Serialize(cellType.ShortcutKeys)));
            result.Add(new JProperty(nameof(cellType.ShowRecommendedValue), cellType.ShowRecommendedValue));
            result.Add(new JProperty(nameof(cellType.ShowTouchToolBar), cellType.ShowTouchToolBar));
            result.Add(new JProperty(nameof(cellType.SideButtons), SideButtonCollectionInfoConverter.Serialize(cellType.SideButtons)));
            result.Add(new JProperty(nameof(cellType.Spin), NumberSpinConverter.Serialize(cellType.Spin)));
            result.Add(new JProperty(nameof(cellType.Static), cellType.Static));
            result.Add(new JProperty(nameof(cellType.TouchContextMenuScale), cellType.TouchContextMenuScale));
            result.Add(new JProperty(nameof(cellType.UseNegativeColor), cellType.UseNegativeColor));
            result.Add(new JProperty(nameof(cellType.UseSpreadDropDownButtonRender), cellType.UseSpreadDropDownButtonRender));
            result.Add(new JProperty(nameof(cellType.ValueSign), cellType.ValueSign));

            return Newtonsoft.Json.JsonConvert.SerializeObject(result);
        }

        /// <summary>
        /// GcNumberCellType をJSONデシリアライズします。
        /// SubEditor, TouchToolBar はデシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">GcNumberCellType オブジェクト。</param>
        /// <param name="cellTypeProps">CellType のプロパティのJSON文字列。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        public static void DeserializeJson(this GcNumberCellType cellType, string cellTypeProps)
        {
            if (string.IsNullOrEmpty(cellTypeProps))
            {
                return;
            }

            var props = JObject.Parse(cellTypeProps);
            foreach (var prop in props)
            {
                if (string.Compare(prop.Key, nameof(GcNumberCellType.AcceptsArrowKeys), true) == 0)
                {
                    cellType.AcceptsArrowKeys = prop.Value.ToObject<AcceptsArrowKeys>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.AcceptsCrLf), true) == 0)
                {
                    cellType.AcceptsCrLf = prop.Value.ToObject<CrLfMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.AcceptsDecimal), true) == 0)
                {
                    cellType.AcceptsDecimal = prop.Value.ToObject<DecimalMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.AllowDeleteToNull), true) == 0)
                {
                    cellType.AllowDeleteToNull = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.AlternateText), true) == 0)
                {
                    NumberAlternateTextInfoConverter.Deserialize(cellType.AlternateText, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.BackgroundImage), true) == 0)
                {
                    cellType.BackgroundImage = PictureConverter.Deserialize(prop.Value, cellType.BackgroundImage);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.ClipContent), true) == 0)
                {
                    cellType.ClipContent = prop.Value.ToObject<ClipContent>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.DisplayFields), true) == 0)
                {
                    cellType.DisplayFields = NumberDisplayFieldCollectionInfoConverter.Deserialize((JArray)prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.DropDown), true) == 0)
                {
                    DropDownInfoConverter.Deserialize(cellType.DropDown, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.DropDownCalculator), true) == 0)
                {
                    DropDownCalculatorInfoConverter.Deserialize(cellType.DropDownCalculator, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.EditMode), true) == 0)
                {
                    cellType.EditMode = prop.Value.ToObject<EditMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.ExcelExportFormat), true) == 0)
                {
                    cellType.ExcelExportFormat = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.ExitOnLastChar), true) == 0)
                {
                    cellType.ExitOnLastChar = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.Fields), true) == 0)
                {
                    NumberFieldsInfoConverter.Deserialize(cellType.Fields, (JArray)prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.FocusPosition), true) == 0)
                {
                    cellType.FocusPosition = prop.Value.ToObject<EditorBaseFocusCursorPosition>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.MaxMinBehavior), true) == 0)
                {
                    cellType.MaxMinBehavior = prop.Value.ToObject<MaxMinBehavior>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.MaxValue), true) == 0)
                {
                    cellType.MaxValue = prop.Value.ToObject<decimal>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.MinValue), true) == 0)
                {
                    cellType.MinValue = prop.Value.ToObject<decimal>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.NegativeColor), true) == 0)
                {
                    cellType.NegativeColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.PaintByControl), true) == 0)
                {
                    cellType.PaintByControl = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.PopUpCalculator), true) == 0)
                {
                    PopUpCalculatorInfoConverter.Deserialize(cellType.PopUpCalculator, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.ReadOnly), true) == 0)
                {
                    cellType.ReadOnly = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.RecommendedValue), true) == 0)
                {
                    cellType.RecommendedValue = prop.Value.ToObject<decimal>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.RoundPattern), true) == 0)
                {
                    cellType.RoundPattern = prop.Value.ToObject<RoundPattern>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.ShortcutKeys), true) == 0)
                {
                    cellType.ShortcutKeys.Clear();
                    ShortcutDictionaryConverter.Deserialize(cellType.ShortcutKeys, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.ShowRecommendedValue), true) == 0)
                {
                    cellType.ShowRecommendedValue = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.ShowTouchToolBar), true) == 0)
                {
                    cellType.ShowTouchToolBar = prop.Value.ToObject<TouchToolBarDisplayOptions>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.SideButtons), true) == 0)
                {
                    cellType.SideButtons = SideButtonCollectionInfoConverter.Deserialize((JArray)prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.Spin), true) == 0)
                {
                    NumberSpinConverter.Deserialize(cellType.Spin, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.Static), true) == 0)
                {
                    cellType.Static = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.TouchContextMenuScale), true) == 0)
                {
                    cellType.TouchContextMenuScale = prop.Value.ToObject<float>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.UseNegativeColor), true) == 0)
                {
                    cellType.UseNegativeColor = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.UseSpreadDropDownButtonRender), true) == 0)
                {
                    cellType.UseSpreadDropDownButtonRender = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcNumberCellType.ValueSign), true) == 0)
                {
                    cellType.ValueSign = prop.Value.ToObject<ValueSignControl>();
                    continue;
                }
            }
        }
    }
}
