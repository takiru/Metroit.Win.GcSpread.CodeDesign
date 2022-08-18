using FarPoint.Win.SuperEdit;
using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.Converters;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// InputManCellTypeBase のコンバーターを提供します。
    /// </summary>
    internal class InputManCellTypeBaseConverter : BaseCellTypeConverter
    {
        /// <summary>
        /// 新しい InputManCellTypeBaseConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public InputManCellTypeBaseConverter(InputManCellTypeBase cellType) : base(cellType)
        {

        }

        /// <summary>
        /// セルタイプのプロパティをシリアライズします。
        /// TouchToolBar はシリアライズを行いません。
        /// </summary>
        /// <param name="propsObj">プロパティオブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        // NOTE: 下記はシリアライズ化対象外。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        protected override void SerializeProp(JObject propsObj, string[] includeProps)
        {
            base.SerializeProp(propsObj, includeProps);

            var c = CellType as InputManCellTypeBase;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.AcceptsArrowKeys))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.AcceptsArrowKeys), c.AcceptsArrowKeys));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.BackgroundImage))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.BackgroundImage), PictureConverter.Serialize(c.BackgroundImage)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.EditMode))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.EditMode), c.EditMode));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.ExcelExportFormat))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.ExcelExportFormat), c.ExcelExportFormat));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.ExitOnLastChar))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.ExitOnLastChar), c.ExitOnLastChar));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.ReadOnly))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.ReadOnly), c.ReadOnly));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.ShortcutKeys))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.ShortcutKeys), ShortcutDictionaryConverter.Serialize(c.ShortcutKeys)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.ShowTouchToolBar))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.ShowTouchToolBar), c.ShowTouchToolBar));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.Static))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.Static), c.Static));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.TouchContextMenuScale))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.TouchContextMenuScale), c.TouchContextMenuScale));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(InputManCellTypeBase.UseSpreadDropDownButtonRender))))
            {
                propsObj.Add(new JProperty(nameof(InputManCellTypeBase.UseSpreadDropDownButtonRender), c.UseSpreadDropDownButtonRender));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// TouchToolBar はデシリアライズを行いません。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as InputManCellTypeBase;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.AcceptsArrowKeys), true) == 0)
            {
                c.AcceptsArrowKeys = prop.Value.ToObject<AcceptsArrowKeys>();
                return;
            }
            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.BackgroundImage), true) == 0)
            {
                c.BackgroundImage = PictureConverter.Deserialize(prop.Value, c.BackgroundImage);
                return;
            }
            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.EditMode), true) == 0)
            {
                c.EditMode = prop.Value.ToObject<EditMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.ExcelExportFormat), true) == 0)
            {
                c.ExcelExportFormat = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.ExitOnLastChar), true) == 0)
            {
                c.ExitOnLastChar = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.ReadOnly), true) == 0)
            {
                c.ReadOnly = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.ShortcutKeys), true) == 0)
            {
                c.ShortcutKeys.Clear();
                ShortcutDictionaryConverter.Deserialize(c.ShortcutKeys, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.ShowTouchToolBar), true) == 0)
            {
                c.ShowTouchToolBar = prop.Value.ToObject<TouchToolBarDisplayOptions>();
                return;
            }
            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.Static), true) == 0)
            {
                c.Static = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.TouchContextMenuScale), true) == 0)
            {
                c.TouchContextMenuScale = prop.Value.ToObject<float>();
                return;
            }
            if (string.Compare(prop.Key, nameof(InputManCellTypeBase.UseSpreadDropDownButtonRender), true) == 0)
            {
                c.UseSpreadDropDownButtonRender = prop.Value.ToObject<bool>();
                return;
            }
        }
    }
}
