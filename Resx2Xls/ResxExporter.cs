using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Resx2Xls
{
    public class ResxExporter
    {
        private readonly DirectoryInfo resxDirectory;
        private readonly DirectoryInfo xlsDirectory;

        public ResxExporter(DirectoryInfo resxDirectory, DirectoryInfo xlsDirectory) 
        {
            this.resxDirectory = resxDirectory;
            this.xlsDirectory = xlsDirectory;
        }

        public void Export(params string[] cultures)
        {
            if (!xlsDirectory.Exists) return;
            if (!resxDirectory.Exists) return;

            var defaultTranslationsFile = resxDirectory.GetFiles("*.resx").FirstOrDefault(cult => cult.GetCulture() == null);
            var defaultTranslations = defaultTranslationsFile.ReadDefaultResxTranslations();
            Parallel.ForEach(cultures, culture =>
                {
                    var sourceFile = resxDirectory.GetFiles($"*.{culture}.resx")
                        .FirstOrDefault(cult => cult.GetCulture() != null && cult.GetCulture().Name == culture);
                    if (sourceFile == null) throw new FileNotFoundException($"File not found for culture ({culture})");
                    
                    var destFile = new FileInfo(Path.Combine(xlsDirectory.FullName, 
                        Path.ChangeExtension(sourceFile.Name, ".xlsx")));
                    
                    var translations = sourceFile.ReadResxTranslations(defaultTranslations);
                    destFile.WriteXlsTranslations(translations);
                });
        }
    }
}