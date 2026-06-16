using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO;
using System;

namespace ExcelDNAtest.Classi
{
    // Cambia la classe in public. Essendo dentro un XLL non cambia nulla all'esterno, 
    // ma il Source Generator ora può "leggerla" senza restrizioni di sicurezza.
    public class ImpostazioniApplicazione
    {
        // Anche le proprietà devono essere public per essere serializzate correttamente in AOT
        public CultureInfo? Lingua { get; set; }
    }

    // Assicurati che anche tutte le sottoclassi siano public!
    // Esempio:
    // public class ImpostazioniSistemaLicenza { ... }

    [JsonSourceGenerationOptions(WriteIndented = true)]
    [JsonSerializable(typeof(ImpostazioniApplicazione))]
    internal partial class AppSettingsJsonContext : JsonSerializerContext
    {
    }

    internal static class SettingsManager
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelToolset", "settings.json");

        private static ImpostazioniApplicazione? _current;
        // Oggetto di blocco per evitare problemi se due thread salvano contemporaneamente
        private static readonly object LockObject = new object();


        /// <summary>
        /// Punto di accesso globale alle impostazioni correnti.
        /// </summary>
        internal static ImpostazioniApplicazione Current
        {
            get
            {
                // Se non è ancora stata caricata, la carichiamo la prima volta (Lazy Loading)
                if (_current == null)
                {
                    _current = Load();
                }
                return _current;
            }
        }

        // Teniamo il metodo Load privato, non serve che gli altri lo chiamino direttamente
        private static ImpostazioniApplicazione Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new ImpostazioniApplicazione();
                string json = File.ReadAllText(FilePath);
                return JsonSerializer.Deserialize(json, AppSettingsJsonContext.Default.ImpostazioniApplicazione) ?? new ImpostazioniApplicazione();
            }
            catch
            {
                return new ImpostazioniApplicazione();
            }
        }

        /// <summary>
        /// Salva le impostazioni correnti su disco in modo sicuro.
        /// </summary>
        internal static void Save() // Rimosso il parametro passatogli, usa direttamente _current
        {
            if (_current == null) return;

            lock (LockObject) // Il lock impedisce scritture corrotte da thread diversi
            {
                try
                {
                    string? dir = Path.GetDirectoryName(FilePath);
                    if (dir != null && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

                    string json = JsonSerializer.Serialize(_current, AppSettingsJsonContext.Default.ImpostazioniApplicazione);
                    File.WriteAllText(FilePath, json);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Errore salvataggio: {ex.Message}");
                }
            }
        }
    }
}