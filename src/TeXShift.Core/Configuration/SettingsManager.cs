using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace TeXShift.Core.Configuration
{
    /// <summary>
    /// Manages loading and saving of AppSettings to JSON file.
    /// Settings are stored in %APPDATA%\TeXShift\settings.json.
    /// </summary>
    public class SettingsManager
    {
        private static readonly string SettingsFolder;
        private static readonly string SettingsFilePath;

        static SettingsManager()
        {
            SettingsFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "TeXShift");
            SettingsFilePath = Path.Combine(SettingsFolder, "settings.json");
        }

        /// <summary>
        /// Gets the path to the settings file.
        /// </summary>
        public static string FilePath => SettingsFilePath;

        /// <summary>
        /// Loads settings from the JSON file.
        /// If the file doesn't exist or is corrupted, returns default settings.
        /// </summary>
        public AppSettings Load()
        {
            try
            {
                if (!File.Exists(SettingsFilePath))
                {
                    return AppSettings.CreateDefault();
                }

                var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                using (var stream = File.OpenRead(SettingsFilePath))
                {
                    var settings = (AppSettings)serializer.ReadObject(stream);
                    return settings ?? AppSettings.CreateDefault();
                }
            }
            catch (Exception)
            {
                // If loading fails for any reason, return defaults
                return AppSettings.CreateDefault();
            }
        }

        /// <summary>
        /// Saves settings to the JSON file.
        /// Creates the settings folder if it doesn't exist.
        /// </summary>
        public void Save(AppSettings settings)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            try
            {
                // Ensure directory exists
                if (!Directory.Exists(SettingsFolder))
                {
                    Directory.CreateDirectory(SettingsFolder);
                }

                var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                using (var stream = File.Create(SettingsFilePath))
                {
                    using (var writer = JsonReaderWriterFactory.CreateJsonWriter(stream, Encoding.UTF8, true, true, "  "))
                    {
                        serializer.WriteObject(writer, settings);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but don't crash
                System.Diagnostics.Debug.WriteLine($"Failed to save settings: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Resets settings to defaults by deleting the settings file.
        /// </summary>
        public void ResetToDefaults()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    File.Delete(SettingsFilePath);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to delete settings file: {ex.Message}");
            }
        }

        /// <summary>
        /// Checks if a settings file exists.
        /// </summary>
        public bool SettingsFileExists()
        {
            return File.Exists(SettingsFilePath);
        }
    }
}
