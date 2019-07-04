﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTranslator.interfaces.FileServices
{
    interface ICanWriteFile
    {
        bool WriteFile(string path, object[] data);
    }
}
