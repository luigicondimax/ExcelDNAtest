using Avalonia.Controls;
using Avalonia.Interactivity;
using ExcelDna.Integration;
using ExcelDNAtest.Classi;
using ExcelDNAtest.Risorse;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace ExcelDNAtest
{
    public partial class FormImpostazioniExcelToolset : Window
    {
        private static ImpostazioniApplicazione _impostazioni => SettingsManager.Current;

        public FormImpostazioniExcelToolset()
        {
            InitializeComponent();

            ApplicaTestiLingua();
            InizializzaDati();
        }

        private void ApplicaTestiLingua()
        {
            Title = "Excel Toolset"; 

            TabGenerali.Header = "Generali";
            LabelLingua.Text = "Seleziona la lingua dell'interfaccia:";

            ButtonSalva.Content = "Salva"; 
            ButtonAnnulla.Content = "Annulla";

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

        internal static void InizializzaComboBoxLingua(Avalonia.Controls.ComboBox comboBox)
        {
            comboBox.ItemsSource = new List<VoceLingua>
                {
                    new() { Codice = "en", Etichetta = "EN - English" },
                    new() { Codice = "it", Etichetta = "IT - Italiano" }
                };


            ImpostazioniApplicazione imp = SettingsManager.Current;

            comboBox.SelectedIndex = (imp.Lingua ?? Thread.CurrentThread.CurrentUICulture).Name switch
            {
                "en" => 0,
                "it" => 1,
                _ => 0,
            };
        }
        private async void ButtonSalva_Click(object? sender, RoutedEventArgs e)
        {
            _impostazioni.Lingua = ComboBoxLingua.SelectedItem is KeyValuePair<CultureInfo, string> kv ? kv.Key : null;
            SettingsManager.Save();

            // RibbonController.myRibbon.Invalidate();

            var box = MessageBoxManager.GetMessageBoxStandard(ImpostazioniRisorse.Lingua_aggiornata, ImpostazioniRisorse.Riavviare_Excel_per_aggiornare_la_lingua_delle_funzioni, ButtonEnum.Ok,MsBox.Avalonia.Enums.Icon.Info);
            await box.ShowWindowAsync();


            this.Close();
        }

        private void ButtonAnnulla_Click(object? sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}