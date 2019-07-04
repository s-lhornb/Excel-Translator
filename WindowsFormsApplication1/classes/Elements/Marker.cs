using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTranslator.classes.Elements
{
    class Marker
    {
        public string name;
        public int number;
        public Marker(string name, int number)
        {
            this.name = name;
            this.number = number;
        }
    }
}
