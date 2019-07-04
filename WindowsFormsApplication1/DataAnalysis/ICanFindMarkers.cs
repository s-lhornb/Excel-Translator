using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTranslator.Elements;

namespace ExcelTranslator.DataAnalysis
{
    interface ICanFindMarkers
    {
        List<Marker> FindMarkers(object[,] table);
    }
}
