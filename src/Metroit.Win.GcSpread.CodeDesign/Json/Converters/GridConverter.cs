using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// Grid オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class GridConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="grid">Grid オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(Grid grid)
        {
            if (grid == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(Grid.Bottom), LineConverter.Serialize(grid.Bottom)));
            jObj.Add(new JProperty(nameof(Grid.Horizontal), LineConverter.Serialize(grid.Horizontal)));
            jObj.Add(new JProperty(nameof(Grid.Left), LineConverter.Serialize(grid.Left)));
            jObj.Add(new JProperty(nameof(Grid.Right), LineConverter.Serialize(grid.Right)));
            jObj.Add(new JProperty(nameof(Grid.Separator), LineConverter.Serialize(grid.Separator)));
            jObj.Add(new JProperty(nameof(Grid.Top), LineConverter.Serialize(grid.Top)));
            jObj.Add(new JProperty(nameof(Grid.Vertical), LineConverter.Serialize(grid.Vertical)));
            jObj.Add(new JProperty(nameof(Grid.VerticalFlag), grid.VerticalFlag));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>Style オブジェクト。</returns>
        public static Grid Deserialize(JToken prop)
        {
            return Deserialize(prop, new Grid());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする Grid オブジェクト。</param>
        /// <returns>Grid オブジェクト。</returns>
        public static Grid Deserialize(JToken prop, Grid source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new Grid();
            if (source != null)
            {
                result.Bottom = source.Bottom;
                result.Horizontal = source.Horizontal;
                result.Left = source.Left;
                result.Right = source.Right;
                result.Separator = source.Separator;
                result.Top = source.Top;
                result.Vertical = source.Vertical;
                result.VerticalFlag = source.VerticalFlag;
            }

            var bottom = prop.SelectToken(nameof(Grid.Bottom));
            if (bottom != null)
            {
                result.Bottom = LineConverter.Deserialize(bottom, result.Bottom);
            }

            var horizontal = prop.SelectToken(nameof(Grid.Horizontal));
            if (horizontal != null)
            {
                result.Horizontal = LineConverter.Deserialize(horizontal, result.Horizontal);
            }

            var left = prop.SelectToken(nameof(Grid.Left));
            if (left != null)
            {
                result.Left = LineConverter.Deserialize(left, result.Left);
            }

            var right = prop.SelectToken(nameof(Grid.Right));
            if (right != null)
            {
                result.Right = LineConverter.Deserialize(right, result.Right);
            }

            var separator = prop.SelectToken(nameof(Grid.Separator));
            if (separator != null)
            {
                result.Separator = LineConverter.Deserialize(separator, result.Separator);
            }

            var top = prop.SelectToken(nameof(Grid.Top));
            if (top != null)
            {
                result.Top = LineConverter.Deserialize(top, result.Top);
            }

            var vertical = prop.SelectToken(nameof(Grid.Vertical));
            if (vertical != null)
            {
                result.Vertical = LineConverter.Deserialize(vertical, result.Vertical);
            }

            var verticalFlag = prop.SelectToken(nameof(Grid.VerticalFlag));
            if (verticalFlag != null)
            {
                result.VerticalFlag = verticalFlag.ToObject<VerticalFlags>();
            }

            return result;
        }
    }
}
