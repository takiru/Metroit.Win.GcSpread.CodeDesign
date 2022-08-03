﻿using FarPoint.Win.Spread;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成する列要素を提供します。
    /// </summary>
    [JsonObject]
    public class ColumnDefinitions
    {
        /// <summary>
        /// 利用するテンプレート名を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string BaseTemplate { get; set; }

        /// <summary>
        /// データバインドするフィールド名を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string DataField { get; set; }

        /// <summary>
        /// 水平位置を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public CellHorizontalAlignment? HorizontalAlignment { get; set; }

        /// <summary>
        /// 垂直位置を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public CellVerticalAlignment? VerticalAlignment { get; set; }

        /// <summary>
        /// 自動フィルター機能を利用するかどうかを取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public bool? AllowAutoFilter { get; set; }

        /// <summary>
        /// 自動ソート機能を利用するかどうかを取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public bool? AllowAutoSort { get; set; }

        /// <summary>
        /// 幅を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public float? Width { get; set; }

        /// <summary>
        /// 表示状態を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public bool? Visible { get; set; }

        /// <summary>
        /// セルタイプを取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public string CellType { get; set; }

        /// <summary>
        /// セルタイプのプロパティを取得または設定します。
        /// サポートするプロパティはセルタイプによって異なります。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public object CellTypeProps { get; set; }

        /// <summary>
        /// IMEモードを取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public System.Windows.Forms.ImeMode? ImeMode { get; set; }

        /// <summary>
        /// ロック状態を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public bool? Locked { get; set; }

        /// <summary>
        /// 外部オプション情報を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public Dictionary<string, object> Options { get; set; }
    }
}
