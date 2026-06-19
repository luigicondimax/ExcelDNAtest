using Avalonia;
using Avalonia.Markup.Xaml;
using Avalonia.Styling;
using Avalonia.Themes.Fluent;
using Avalonia.Threading;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
#if !AOT
using ExcelDna.IntelliSense;
#endif
using ExcelDna.Logging;
using ExcelDna.Registration;
using ExcelDNAtest.Classi;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;

namespace ExcelDNAtest
{
    public partial class Events : IExcelAddIn
    {

        internal static bool MostraRibbon = false;
        private static AppBuilder? _appBuilder;
        private static ApplicationSettings imp => SettingsManager.Current;

        public void AddIn()
        {

        }

        public async void AutoOpen()
        {
            Trace.Listeners.Add(new LogDisplayTraceListener());

            ParameterConversionConfiguration conversionConfig;
            conversionConfig = new ParameterConversionConfiguration().AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true));

            if (imp.Language is not null)
            {
                var culture = CultureInfo.GetCultureInfoByIetfLanguageTag(imp.Language);
                CultureInfo.CurrentCulture = culture;
                CultureInfo.CurrentUICulture = culture;
                CultureInfo.DefaultThreadCurrentCulture = culture;
                CultureInfo.DefaultThreadCurrentUICulture = culture;
            }
            Trace.TraceInformation($"CurrentUICulture: {CultureInfo.CurrentUICulture}");

            try
            {
#if AOT
                //		<NativeLibrary Include="av_libglesv2.dll" />

                string xllDirectory = Path.GetDirectoryName(ExcelDnaUtil.XllPath)!;
                string skiaDllPath = Path.Combine(xllDirectory, "libSkiaSharp.dll");
                string avDllPath = Path.Combine(xllDirectory, "av_libglesv2.dll");
                string harfBuzzDllPath = Path.Combine(xllDirectory, "libHarfBuzzSharp.dll"); // <-- NUOVA

                // Forza il caricamento in memoria di libSkiaSharp.dll
                if (File.Exists(skiaDllPath))
                {
                    NativeLibrary.Load(skiaDllPath);
                }
                else
                {
                    Trace.TraceError($"Attenzione: libSkiaSharp.dll non trovata in {skiaDllPath}");
                }

                // Forza il caricamento in memoria di av_libglesv2.dll
                if (File.Exists(avDllPath))
                {
                    NativeLibrary.Load(avDllPath);
                }
                else
                {
                    Trace.TraceError($"Attenzione: av_libglesv2.dll non trovata in {avDllPath}");
                }

                // NUOVO: Forza il caricamento in memoria di libHarfBuzzSharp.dll prima del rendering del testo
                if (File.Exists(harfBuzzDllPath))
                {
                    NativeLibrary.Load(harfBuzzDllPath);
                }
                else
                {
                    Trace.TraceError($"Attenzione: libHarfBuzzSharp.dll non trovata in {harfBuzzDllPath}");
                } 
#endif

                _appBuilder = AppBuilder.Configure<App>()
                    .UsePlatformDetect()
                    .LogToTrace()
                    .SetupWithoutStarting();
            }
            catch (Exception e)
            {
                Trace.TraceError(e.ToString());
            }



            DynamicRegistration.RegisterLocalizedFunctions();

#if !AOT
            IntelliSenseServer.Install();
#endif

            CreaContextMenu();

        }

        private static void CreaContextMenu()
        {
            // Inserisco nel menu contestuale tasto destro alcune nuove macro
            CommandBars bars = ExcelCommandBarUtil.GetCommandBars();

            CommandBar menuBar = bars["Cell"];


            CommandBarButton c = menuBar.Controls.AddButton();
            c.Caption = "Test";
            c.Style = MsoButtonStyle.msoButtonCaption;
            c.OnAction = "CreateMsg";
        }
        public void AutoClose()
        {

#if !AOT
            IntelliSenseServer.Uninstall();
#endif
        }
        [ExcelCommand]
        public static async void CreateMsg()
        {
            await Dispatcher.UIThread.InvokeAsync(async () =>
            {
                try
                {
                    var box = MessageBoxManager.GetMessageBoxStandard("ContextMenu", "Clicked", ButtonEnum.Ok, Icon.Warning);

                    await box.ShowWindowAsync();
                }
                catch (Exception e)
                {

                    Trace.TraceError($"Errore in CreateMessageBox: {e.Message}");
                    throw;
                }
            });
        }

    }
    public class App : Avalonia.Application
    {
        public App()
        {
            Styles.Add(new FluentTheme());
            RequestedThemeVariant = ThemeVariant.Light; // oppure Dark, oppure Default
        }
    }
}
