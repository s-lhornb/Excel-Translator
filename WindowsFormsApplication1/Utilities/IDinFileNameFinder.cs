using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTranslator.Utilities;

namespace ExcelTranslator.Utilities
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
