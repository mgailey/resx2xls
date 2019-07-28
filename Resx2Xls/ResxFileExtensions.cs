using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Resources;

namespace Resx2Xls
{
	public static class ResxFileExtensions
	{
		public static IReadOnlyCollection<Translation> ReadDefaultResxTranslations(this FileInfo sourceFile)
		{
			var rows = new List<Translation>();
			using (var reader = new ResXResourceReader(sourceFile.FullName))
			{
				reader.UseResXDataNodes = true;
				foreach (DictionaryEntry de in reader)
				{
					if (!(de.Value is ResXDataNode dataNode)) continue;

					var key = (string)de.Key;
					var value = (string)dataNode.GetValue((ITypeResolutionService)null);
					rows.Add(new Translation
					{
						Key = key,
						Source = value,
						Usage = dataNode.Comment
					});
				}
			}
			return rows;
		}
		
		public static IReadOnlyCollection<Translation> ReadResxTranslations(this FileInfo sourceFile, IReadOnlyCollection<Translation> defaultTranslations)
		{
			var rows = new List<Translation>();
			using (var reader = new ResXResourceReader(sourceFile.FullName))
			{
				foreach (DictionaryEntry de in reader)
				{
					if (!(de.Value is string)) continue;

					var key = (string)de.Key;

					var row = defaultTranslations.FirstOrDefault(tr => tr.Key == key);
					if (row == null) throw new Exception("Row not Found");

					var clone = (Translation) row.Clone();
					clone.Native = de.Value.ToString();
					rows.Add(clone);
				}
			}
			return rows;
		}

		public static void WriteResxTranslations(this FileInfo destFile, IReadOnlyCollection<Translation> translations)
		{
			using (var rw = new ResXResourceWriter(destFile.FullName))
			{
				foreach (var translation in translations.Where(tr => tr.HasTranslation))
				{
					var resXDataNode = new ResXDataNode(translation.Key, translation.Native);
					rw.AddResource(resXDataNode);
				}
			}
		}
	}
}