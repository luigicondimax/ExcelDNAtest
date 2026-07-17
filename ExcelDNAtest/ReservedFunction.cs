using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDNAtest
{
    internal class ReservedFunction
    {
        [ExcelFunction(ExplicitRegistration = true)]
        internal static string DynamicRegistrationTest()
        {
            return "Hello world from reserved!";
        }
    }
}
