using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelDNAtest.Classi
{
    // Cambia la classe in public. Essendo dentro un XLL non cambia nulla all'esterno, 
    // ma il Source Generator ora può "leggerla" senza restrizioni di sicurezza.
    public class ApplicationSettings
    {
        // Anche le proprietà devono essere public per essere serializzate correttamente in AOT
        public string? Language { get; set; }
    }

    // Assicurati che anche tutte le sottoclassi siano public!
    // Esempio:
    // public class ImpostazioniSistemaLicenza { ... }

    [JsonSourceGenerationOptions(WriteIndented = true)]
    [JsonSerializable(typeof(ApplicationSettings))]
    internal partial class AppSettingsJsonContext : JsonSerializerContext
    {
    }

    internal static class SettingsManager
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelToolset", "settings.json");

        private static ApplicationSettings? _current;
        // Oggetto di blocco per evitare problemi se due thread salvano contemporaneamente
        private static readonly Lock LockObject = new();


        /// <summary>
        /// Punto di accesso globale alle impostazioni correnti.
        /// </summary>
        internal static ApplicationSettings Current
        {
            get
            {
                // Se non è ancora stata caricata, la carichiamo la prima volta (Lazy Loading)
                _current ??= Load();
                return _current;
            }
        }

        // Teniamo il metodo Load privato, non serve che gli altri lo chiamino direttamente
        private static ApplicationSettings Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new ApplicationSettings();
                string json = File.ReadAllText(FilePath);
                return JsonSerializer.Deserialize(json, AppSettingsJsonContext.Default.ApplicationSettings) ?? new ApplicationSettings();
            }
            catch
            {
                return new ApplicationSettings();
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

                    string json = JsonSerializer.Serialize(_current, AppSettingsJsonContext.Default.ApplicationSettings);
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