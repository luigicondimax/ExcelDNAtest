using ExcelDna.Integration;
using System.Globalization;
using System.Text.RegularExpressions;
using static ExcelDNAtest.ExcelUtil;

namespace ExcelDNAtest
{
    public class Function
    {
        [ExcelFunction(description:"Descrizione",ExplicitRegistration =true)]
        public static string Hello()
        {
            return "Hello world!";
        }

#if AOT
        [ExcelFunction]
        public static double NativeApplicationGetCellValue(string cell)
        {
            var workbook = (IDynamic)ExcelDnaUtil.DynamicApplication.Get("ActiveWorkbook")!;
            var sheets = (IDynamic)workbook.Get("Sheets")!;
            var sheet = (IDynamic)sheets[1]!;
            var range = (IDynamic)sheet.Get("Range", [cell])!;
            return (double)range.Get("Value")!;
        }

        [ExcelFunction]
        public static string NativeApplicationGetCellFormula(string cell)
        {
            // in AOT return local language formula, in JIT return English formula
            // AOT use the excel language and not the CultureInfo.

            //var culture = CultureInfo.GetCultureInfo("en-US");

            //CultureInfo.CurrentCulture = culture;
            //CultureInfo.CurrentUICulture = culture;
            //CultureInfo.DefaultThreadCurrentCulture = culture;
            //CultureInfo.DefaultThreadCurrentUICulture = culture;

            var workbook = (IDynamic)ExcelDnaUtil.DynamicApplication.Get("ActiveWorkbook")!;
            var sheets = (IDynamic)workbook.Get("Sheets")!;
            var sheet = (IDynamic)sheets[1]!;
            var range = (IDynamic)sheet.Get("Range", [cell])!;
            return (string)range.Get("Formula")!;
        }
#else
        [ExcelFunction(ExplicitRegistration = true)]
        public static double NativeApplicationGetCellValue(string cell)
        {
            var workbook = ExcelUtil.ActiveWorkbook();
            var sheet = workbook!.Sheets[1];
            var range = sheet.Range[cell];
            return (double)range.Value;
        }
        [ExcelFunction( ExplicitRegistration = true)]
        public static string NativeApplicationGetCellFormula(string cell)
        {
            var workbook = ExcelUtil.ActiveWorkbook()!;
            var sheet = workbook.Sheets[1];
            var range = sheet.Range[cell];
            return (string)range.Formula;
        }
#endif

    }

}
