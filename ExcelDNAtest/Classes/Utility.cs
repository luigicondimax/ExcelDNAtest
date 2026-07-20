using ExcelDNAtest.Classi;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelDNAtest.Classes
{
    internal static class Utility
    {
        private static ApplicationSettings imp => SettingsManager.Current;

        internal static void ChangeLanguage()
        {
            if (imp.Language is not null)
            {
                var culture = CultureInfo.GetCultureInfoByIetfLanguageTag(imp.Language);
                CultureInfo.CurrentCulture = culture;
                CultureInfo.CurrentUICulture = culture;
                CultureInfo.DefaultThreadCurrentCulture = culture;
                CultureInfo.DefaultThreadCurrentUICulture = culture;
            }
        }
    }
}
