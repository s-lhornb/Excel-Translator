using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTranslator.Elements;

namespace ExcelTranslator.DataAnalysis
{
    interface ICanBuildMarkerPairs
    {
        List<Marker[]> FindMarkerPairs(List<Marker> markers);
    }
}
