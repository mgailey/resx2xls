using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Resx2Xls
{
    public class ResxImporter
    {
        private readonly DirectoryInfo resxDirectory;
        private readonly DirectoryInfo xlsDirectory;

        public ResxImporter(DirectoryInfo resxDirectory, DirectoryInfo xlsDirectory)
        {
            this.resxDirectory = resxDirectory;
            this.xlsDirectory = xlsDirectory;
        }

        public void Import(params string[] cultures)
        {
            if (!resxDirectory.Exists) return;
            if (!xlsDirectory.Exists) return;
            
            Parallel.ForEach(cultures, culture =>
                {
                    var sourceFile = xlsDirectory.GetFiles($"*.{culture}.xlsx")
                        .FirstOrDefault(cult => cult.GetCulture() != null && cult.GetCulture().Name == culture);
                    if (sourceFile == null) throw new FileNotFoundException($"File not found for culture ({culture})");
                    
                    var destFile = new FileInfo(Path.Combine(resxDirectory.FullName, 
                        Path.ChangeExtension(sourceFile.Name, ".resx")));
                    
                    var translations = sourceFile.ReadXlsTranslations();
                    destFile.WriteResxTranslations(translations);
                });
        }
    }
}