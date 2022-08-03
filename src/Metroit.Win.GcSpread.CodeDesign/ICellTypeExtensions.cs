using FarPoint.Win.Spread.CellType;
using GrapeCity.Win.Spread.InputMan.CellType;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Metroit.Win.GcSpread.CodeDesign
{
    internal static class ICellTypeExtensions
    {
        public static void DeserializeJson(this ICellType cellType, string cellTypeProps)
        {
            if (cellType is GcNumberCellType)
            {
                ((GcNumberCellType)cellType).DeserializeJson(cellTypeProps);
            }
        }
    }
}
