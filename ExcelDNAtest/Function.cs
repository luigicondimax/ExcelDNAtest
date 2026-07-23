using Avalonia.Threading;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using static ExcelDNAtest.ExcelUtil;
using System;

namespace ExcelDNAtest
{
    public class Function
    {
        [ExcelFunction(description: "Descrizione", ExplicitRegistration = true)]
        public static string Hello()
        {
            return "Hello world!";
        }

        [ExcelFunction(IsMacroType = true)]

        public static bool LanguageTest(object _true, object _sum)
        {
            // =LanguageTest(VERO;SOMMA(1;2))
            // format #.##0,00


            // IN AOT always local language

            // C-API
            var caller = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            Trace.TraceInformation($"C-API Local formula: {XlCall.Excel(XlCall.xlfGetCell, 6, caller)}");
            Trace.TraceInformation($"C-API Global formula: {XlCall.Excel(XlCall.xlfGetCell, 41, caller)}");
            Trace.TraceInformation($"C-API Format: {XlCall.Excel(XlCall.xlfGetCell, 7, caller)}");

            // COM
#if AOT
            var range = ToRange(caller)!;

            Trace.TraceInformation($"COM Formula: {range.Get("Formula")}");
            Trace.TraceInformation($"COM Formula: {range.Get("FormulaLocal")}");
            Trace.TraceInformation($"COM Format: {range.Get("NumberFormat")}");
#else
            var range = ToRange(caller)!;

            Trace.TraceInformation($"COM Formula: {range.Formula}");
            Trace.TraceInformation($"COM Formula: {range.FormulaLocal}");
            Trace.TraceInformation($"COM Format: {range.NumberFormat}");
#endif
            // COM LCID english
#if AOT
            NativeLocaleScope.ExecuteInEnglish(() =>
            {
                var range = ToRange(caller)!;

                Trace.TraceInformation($"COM LCID Formula: {range.Get("Formula")}");
                Trace.TraceInformation($"COM LCID Formula: {range.Get("FormulaLocal")}");
                Trace.TraceInformation($"COM LCID Format: {range.Get("NumberFormat")}");
            });
#else
            NativeLocaleScope.ExecuteInEnglish(() =>
            {
                Trace.TraceInformation($"COM LCID Formula: {range.Formula}");
                Trace.TraceInformation($"COM LCID Formula: {range.FormulaLocal}");
                Trace.TraceInformation($"COM LCID Format: {range.NumberFormat}");
            });

#endif
            // JIT result
            //Information: 0 : C - API Local formula: = LanguageTest(VERO; SOMMA(1; 2))
            //Information: 0 : C - API Global formula: = LanguageTest(VERO; SOMMA(1; 2))
            //Information: 0 : C - API Format: #.##0,00
            //Information: 0 : COM Formula: = LanguageTest(TRUE, SUM(1, 2))
            //Information: 0 : COM Formula: = LanguageTest(VERO; SOMMA(1; 2))
            //Information: 0 : COM Format: #,##0.00
            //Information: 0 : COM LCID Formula: = LanguageTest(TRUE, SUM(1, 2))
            //Information: 0 : COM LCID Formula: = LanguageTest(VERO; SOMMA(1; 2))
            //Information: 0 : COM LCID Format: #,##0.00

            // AOT
            //Information: 0 : C - API Local formula: = LanguageTest(VERO; SOMMA(1; 2))
            //Information: 0 : C - API Global formula: = LanguageTest(VERO; SOMMA(1; 2))
            //Information: 0 : C - API Format: #.##0,00
            //Information: 0 : COM Formula: = LanguageTest(VERO; SOMMA(1; 2))
            //Information: 0 : COM Formula: = LanguageTest(VERO; SOMMA(1; 2))
            //Information: 0 : COM Format: #.##0,00
            //Information: 0 : COM LCID Formula: = LanguageTest(VERO; SOMMA(1; 2))
            //Information: 0 : COM LCID Formula: = LanguageTest(VERO; SOMMA(1; 2))
            //Information: 0 : COM LCID Format: #.##0,00

            return true;
        }
    }

    public static class NativeLocaleScope
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetThreadLocale(uint Lcid);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint GetThreadLocale();

        public const uint LCID_EN_US = 1033; // 0x0409 (en-US)
        private static readonly CultureInfo CultureEnUs = CultureInfo.GetCultureInfo("en-US");
        /// <summary>
        /// Esegue un blocco di codice (Action) forzando l'LCID nativo del thread a en-US (1033).
        /// </summary>
        public static void ExecuteInEnglish(System.Action action)
        {
            uint originalLcid = GetThreadLocale();
            var originalCulture = CultureInfo.CurrentCulture;
            var originalUiCulture = CultureInfo.CurrentUICulture;

            try
            {
                // 1. Cambia il thread nativo Win32
                SetThreadLocale(LCID_EN_US);

                // 2. Cambia la cultura gestita .NET (se IDynamic la legge per impostare lcid)
                CultureInfo.CurrentCulture = CultureEnUs;
                CultureInfo.CurrentUICulture = CultureEnUs;

                action();
            }
            finally
            {
                SetThreadLocale(originalLcid);
                CultureInfo.CurrentCulture = originalCulture;
                CultureInfo.CurrentUICulture = originalUiCulture;
            }
        }
    }

}
