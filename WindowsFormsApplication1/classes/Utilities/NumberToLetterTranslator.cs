using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTranslator.interfaces.Utilities;

namespace ExcelTranslator.classes.Utilities
{
    class NumberToLetterTranslator : ICanTranslateNumbersToLetters
    {
        public string TranslaNumber(int number)
        {
            string translated = String.Empty;
            int modulo;

            while (number > 0)
            {
                modulo = (number - 1) % 26;
                translated = Convert.ToChar(modulo + 65).ToString() + translated;
                number = (int)((number - modulo) / 26);
            }

            return translated;
        }
    }
}
