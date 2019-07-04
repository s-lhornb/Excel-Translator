using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTranslator.classes.Elements;

namespace ExcelTranslator.interfaces.DataAnalysis
{
    interface ICanBuildMarkerPairs
    {
        List<Marker[]> FindMarkerPairs(List<Marker> markers);
    }
}
