using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Resources;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ExcelDNAtest
{
    internal static partial class DynamicRegistration
    {
        // Ensures AOT compiler doesn't trim away the functions or the resource file
        [DynamicDependency(DynamicallyAccessedMemberTypes.PublicMethods, typeof(ReservedFunction))]
        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(Resources.ReservedFunction))]
        internal static void RegisterLocalizedFunctions()
        {
            var resourceManagers = new Dictionary<string, ResourceManager>()
            {
                [nameof(ReservedFunction)] = new ResourceManager(typeof(Resources.ReservedFunction))
            };

            var mainCulture = CultureInfo.CurrentCulture;
            CultureInfo[] secondaryCultures = [mainCulture.TwoLetterISOLanguageName == "en" ? new CultureInfo("it") : new CultureInfo("en")];

            // Optimized Reflection using C# Property Pattern Matching
            var licensedMethods = typeof(ReservedFunction)
                .GetMethods(BindingFlags.Public | BindingFlags.Static)
                .Where(m => m.GetCustomAttribute<ExcelFunctionAttribute>() is { ExplicitRegistration: true })
                .ToList();

            ResourceManager resManager = resourceManagers[nameof(ReservedFunction)];
            var allRegistrations = new List<ExcelFunctionRegistration>();

            // --- 1. Primary Language Registration ---
            var primaryRegistrations = licensedMethods
                .Select(m => new ExcelFunctionRegistration(m))
                .ProcessParamsRegistrations()
                .ToList();

            for (int i = 0; i < primaryRegistrations.Count; i++)
            {
                var funcReg = primaryRegistrations[i];
                var methodInfo = licensedMethods[i];

                bool translationAvailable = LocalizeFunctionRecord(funcReg, methodInfo.Name, mainCulture, resManager, isPrimary: true);

                if (translationAvailable)
                    allRegistrations.Add(funcReg);
                else
                    Debug.Print($"Function '{methodInfo.Name}' lacks translation for primary culture '{mainCulture.TwoLetterISOLanguageName}'. Skipping alias.");
            }

            // --- 2. Secondary Languages (Hidden Aliases) ---
            foreach (var secondaryCulture in secondaryCultures)
            {
                var secondaryRegistrations = licensedMethods
                    .Select(m => new ExcelFunctionRegistration(m))
                    .ProcessParamsRegistrations()
                    .ToList();

                for (int i = 0; i < secondaryRegistrations.Count; i++)
                {
                    var funcReg = secondaryRegistrations[i];
                    var methodInfo = licensedMethods[i];

                    bool translationAvailable = LocalizeFunctionRecord(funcReg, methodInfo.Name, secondaryCulture, resManager, isPrimary: false);

                    // Hide from Function Wizard autocomplete, but keep it functional
                    funcReg.FunctionAttribute.IsHidden = true;

                    if (translationAvailable)
                        allRegistrations.Add(funcReg);
                    else
                        Debug.Print($"Function '{methodInfo.Name}' lacks translation for secondary culture '{secondaryCulture.TwoLetterISOLanguageName}'. Skipping alias.");
                }
            }

            // --- 3. Final Excel-DNA Registration ---
            allRegistrations.RegisterFunctions();
        }

        private static bool LocalizeFunctionRecord(ExcelFunctionRegistration funcReg, string originalMethodName, CultureInfo culture, ResourceManager resManager, bool isPrimary)
        {
            string nameKey = $"{originalMethodName}_Name";
            string descKey = $"{originalMethodName}_Desc";

            string? localizedName = resManager.GetString(nameKey, culture);
            string? localizedDesc = resManager.GetString(descKey, culture);

            if (string.IsNullOrEmpty(localizedName))
            {
                if (isPrimary) localizedName = originalMethodName;
                else return false; // Fail silently for secondary aliases if no translation is found
            }

            if (string.IsNullOrEmpty(localizedDesc) && isPrimary)
            {
                localizedDesc = funcReg.FunctionAttribute.Description;
            }

            funcReg.FunctionAttribute.Name = localizedName;
            funcReg.FunctionAttribute.Description = localizedDesc;

            // --- Argument Localization (Double Underscore Pattern) ---
            for (int i = 0; i < funcReg.ParameterRegistrations.Count; i++)
            {
                var paramReg = funcReg.ParameterRegistrations[i];
                int argumentNumber = i + 1;

                // 1. Specific key search (e.g., MyFunction__arg1_Name)
                string argNameKey = $"{originalMethodName}__arg{argumentNumber}_Name";
                string argDescKey = $"{originalMethodName}__arg{argumentNumber}_Desc";

                string? localizedArgName = resManager.GetString(argNameKey, culture);
                string? localizedArgDesc = resManager.GetString(argDescKey, culture);

                // 2. Fallback: Generic key search (e.g., __arg1_Name)
                if (string.IsNullOrEmpty(localizedArgName))
                {
                    localizedArgName = resManager.GetString($"__arg{argumentNumber}_Name", culture);
                }
                if (string.IsNullOrEmpty(localizedArgDesc))
                {
                    localizedArgDesc = resManager.GetString($"__arg{argumentNumber}_Desc", culture);
                }

                if (!string.IsNullOrEmpty(localizedArgName)) paramReg.ArgumentAttribute.Name = localizedArgName;
                if (!string.IsNullOrEmpty(localizedArgDesc)) paramReg.ArgumentAttribute.Description = localizedArgDesc;
            }

            return true;
        }
    }
}