using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using FrEngineLoader;

namespace Sample
{
    internal class Program
    {
        //[STAThread] // Attribute is required only in case of using FrEngineLoadingMethod.Native.
        private static void Main()
        {
            dynamic engine = null;

            try
            {
                var inputFolderPath = PrepareDirectory(ConfigurationManager.AppSettings["InputFolderPath"] ?? "Input");
                var outputFolderPath = PrepareDirectory(ConfigurationManager.AppSettings["OutputFolderPath"] ?? "Output");
                var logFileName = ConfigurationManager.AppSettings["LogFileName"] ?? "FREngine.log";
                var settingsFileName = ConfigurationManager.AppSettings["SettingsFileName"] ?? string.Empty;
                var executingAssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ??
                                                 string.Empty;

                // Initialize wrapped ABBYY FineReader Engine 11 COM object.
                engine = new DynamicFrEngine(FrEngineLoadingMethod.OutOfProcessComServer);

                // Enable FrEngine logging.
                engine.StartLogging(Path.Combine(executingAssemblyDirectory, logFileName), true);

                // Load FrEngine profile with custom settings.
                engine.LoadProfile(Path.Combine(executingAssemblyDirectory, settingsFileName));

                // Fetch all supported image files from Input folder.
                var imageFiles = (from imageFile in Directory.EnumerateFiles(inputFolderPath)
                                  where DynamicFrEngine.SupportedImageExtensions.Any(
                                    extension => imageFile.ToLower().EndsWith(extension))
                                  select imageFile).ToArray();

                // Finish if no images in Input folder.
                if (imageFiles.Length == 0)
                {
                    Console.WriteLine(@"No images in '{0}'", inputFolderPath);
                    return;
                }

                using (dynamic frDocument = engine.CreateFRDocument())
                {
                    foreach (var imageFile in imageFiles)
                    {
                        frDocument.AddImageFile(imageFile, null, null);
                        frDocument.Process(null);
                        frDocument.Export(Path.Combine(outputFolderPath, Path.GetFileName(imageFile) + ".docx"),
                            FileExportFormat.Docx, null);
                        frDocument.Close();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

                if (e.InnerException != null)
                {
                    Console.WriteLine(@"Inner exception message: {0}", e.InnerException.Message);
                }
            }
            finally
            {
                if (engine != null) engine.Dispose();
            }
        }

        private static string PrepareDirectory(string directory)
        {
            return Directory.CreateDirectory(directory).FullName;
        }
    }
}