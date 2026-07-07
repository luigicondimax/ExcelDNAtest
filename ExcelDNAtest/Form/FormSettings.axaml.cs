using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using ExcelDNAtest.Classi;
using ExcelDNAtest.Resources;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using System.Diagnostics;
using System.Globalization;

namespace ExcelDNAtest
{
    public partial class FormSettings : Window
    {
        private static ApplicationSettings _settings => SettingsManager.Current;

        public FormSettings()
        {
            string? languageSaved = _settings.Language;
            if (!string.IsNullOrEmpty(languageSaved))
            {
                var culture = CultureInfo.GetCultureInfoByIetfLanguageTag(languageSaved);

                CultureInfo.CurrentCulture = culture;
                CultureInfo.CurrentUICulture = culture;

                CultureInfo.DefaultThreadCurrentCulture = culture;
                CultureInfo.DefaultThreadCurrentUICulture = culture;
            }

            InitializeComponent();

            ApplicaTestiLingua();
            InizializzaDati();
        }

        private void ApplicaTestiLingua()
        {
            Title = "Excel test";

            TabGenerali.Header = SettingsResources.Localizzazione;
            LabelLingua.Text = SettingsResources.Seleziona_la_lingua_dell_interfaccia_;

            ButtonSalva.Content = SettingsResources.Salva;
            ButtonAnnulla.Content = SettingsResources.Annulla;

        }

        private void InizializzaDati()
        {
            InizializzaComboBoxLingua(ComboBoxLingua);
        }

        internal sealed class VoceLingua
        {
            public string Codice { get; init; } = "";
            public string Etichetta { get; init; } = "";

            // Questo dice ad Avalonia cosa mostrare graficamente senza usare la riflessione
            public override string ToString() => Etichetta;
        }

        internal static void InizializzaComboBoxLingua(ComboBox comboBox)
        {
            comboBox.ItemsSource = new List<VoceLingua>
                {
                    new() { Codice = "en", Etichetta = "EN - English" },
                    new() { Codice = "it", Etichetta = "IT - Italiano" }
                };


            ApplicationSettings imp = SettingsManager.Current;

            comboBox.SelectedIndex = (imp.Language ?? Thread.CurrentThread.CurrentUICulture.Name) switch
            {
                "en" => 0,
                "it" => 1,
                _ => 0,
            };
        }
        private async void ButtonSalva_Click(object? sender, RoutedEventArgs e)
        {
            if (ComboBoxLingua.SelectedItem is VoceLingua voceSelezionata)
            {
                _settings.Language = voceSelezionata.Codice;
                SettingsManager.Save();

                var culture = CultureInfo.GetCultureInfoByIetfLanguageTag(voceSelezionata.Codice);
                CultureInfo.CurrentCulture = culture;
                CultureInfo.CurrentUICulture = culture;
                CultureInfo.DefaultThreadCurrentCulture = culture;
                CultureInfo.DefaultThreadCurrentUICulture = culture;
            }
            else
            {
                Trace.TraceError("Nessuna lingua selezionata, impostazione non aggiornata.");
            }

            RibbonController.myRibbon?.Invalidate();


            var titolo = SettingsResources.Lingua_aggiornata;
            var testo = SettingsResources.Riavviare_Excel_per_aggiornare_la_lingua_delle_funzioni;
            var box = MessageBoxManager.GetMessageBoxStandard(titolo, testo, ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Info);
            await box.ShowWindowDialogAsync(this);

            // or close before showing the message box.
            Dispatcher.UIThread.Post(() => this.Close());
        }

        private void ButtonAnnulla_Click(object? sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}