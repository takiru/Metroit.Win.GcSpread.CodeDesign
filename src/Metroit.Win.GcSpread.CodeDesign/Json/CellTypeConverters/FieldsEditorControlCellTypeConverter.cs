using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// FieldsEditorControlCellType のコンバーターを提供します。
    /// </summary>
    internal class FieldsEditorControlCellTypeConverter : InputManCellTypeBaseConverter
    {
        /// <summary>
        /// 新しい FieldsEditorControlCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public FieldsEditorControlCellTypeConverter(FieldsEditorControlCellType cellType) : base(cellType)
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

            var c = CellType as FieldsEditorControlCellType;
            if (c == null)
            {
                return;
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(FieldsEditorControlCellType.ClipContent))))
            {
                propsObj.Add(new JProperty(nameof(FieldsEditorControlCellType.ClipContent), c.ClipContent));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(FieldsEditorControlCellType.PromptChar))))
            {
                propsObj.Add(new JProperty(nameof(FieldsEditorControlCellType.PromptChar), c.PromptChar.ToString()));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(FieldsEditorControlCellType.ShowLiterals))))
            {
                propsObj.Add(new JProperty(nameof(FieldsEditorControlCellType.ShowLiterals), c.ShowLiterals));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(FieldsEditorControlCellType.TabAction))))
            {
                propsObj.Add(new JProperty(nameof(FieldsEditorControlCellType.TabAction), c.TabAction));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as FieldsEditorControlCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(FieldsEditorControlCellType.ClipContent), true) == 0)
            {
                c.ClipContent = prop.Value.ToObject<ClipContent>();
                return;
            }
            if (string.Compare(prop.Key, nameof(FieldsEditorControlCellType.PromptChar), true) == 0)
            {
                c.PromptChar = prop.Value.ToObject<char>();
                return;
            }
            if (string.Compare(prop.Key, nameof(FieldsEditorControlCellType.ShowLiterals), true) == 0)
            {
                c.ShowLiterals = prop.Value.ToObject<ShowLiterals>();
                return;
            }
            if (string.Compare(prop.Key, nameof(FieldsEditorControlCellType.TabAction), true) == 0)
            {
                c.TabAction = prop.Value.ToObject<TabAction>();
                return;
            }
        }
    }
}
