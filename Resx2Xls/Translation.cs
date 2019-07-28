using System;

namespace Resx2Xls
{
	public class Translation : ICloneable
	{
		public string Key { get; set; }
		public string Source { get; set; }
		public string Native { get; set; }
		public string Usage { get; set; }

		public bool HasTranslation => !string.IsNullOrEmpty(Key) &&
		                         !string.IsNullOrEmpty(Native);

		public object Clone()
		{
			return new Translation {Key = Key, Source = Source, Native = Native, Usage = Usage};
		}
	}
}