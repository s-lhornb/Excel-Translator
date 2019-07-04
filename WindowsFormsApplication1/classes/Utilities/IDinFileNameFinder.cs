using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTranslator.interfaces.Utilities;

namespace ExcelTranslator.classes.Utilities
{
    class IDinFileNameFinder : ICanFindIDInFileNames
    {
        public string FindID(string path)
        {
            string[] splitData = path.Split('\\');
            return splitData[splitData.Length - 1].Split('_')[0];
        }
    }
}
