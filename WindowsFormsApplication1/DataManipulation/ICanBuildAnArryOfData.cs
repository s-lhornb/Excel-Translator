using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTranslator.DataManipulation
{
    interface ICanBuildAnArryOfData
    {
        object[] buildArray(List<T> data);
    }
}
