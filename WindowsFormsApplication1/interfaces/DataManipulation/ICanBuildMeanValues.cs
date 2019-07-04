using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTranslator.classes.Elements;

namespace ExcelTranslator.interfaces.DataManipulation
{
    interface ICanBuildMeanValues
    {
        List<Tuple<Marker[], List<double>>> buildMeanValue (List<Marker[]> markerpairs, object[,] table);
    }
}
