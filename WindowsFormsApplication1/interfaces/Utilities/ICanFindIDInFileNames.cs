﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTranslator.interfaces.Utilities
{
    interface ICanFindIDInFileNames
    {
        string FindID(string path);
    }
}