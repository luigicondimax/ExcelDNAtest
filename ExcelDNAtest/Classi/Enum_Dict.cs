using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDNAtest.Classi
{
    internal class Enum_Dict
    {
        public enum XlInputBoxDataType
        {
            Formula = 0,
            Number = 1,
            Text = 2,
            Boolean = 4,
            Range = 8,
            Error = 16,
            Array = 64
        }

    }
}
