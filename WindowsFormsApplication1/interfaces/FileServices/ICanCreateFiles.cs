using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTranslator.interfaces.FileServices
{
    interface ICanCreateFile
    {
        bool CreateFile(string path);
    }
}
