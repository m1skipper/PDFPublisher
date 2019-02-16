using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PDFCompare
{
	public class Settings
	{
		[ValueList(typeof(List<string>), MaximumElements = 3)]
		public IList<string> Files { get; set; }
		[Option('c', "changed", DefaultValue = 0x0000FFFF, HelpText = "Цвет пометки изменённых фрагментов")]
		public int ColorChanged { get; set; }
		[Option('a', "added", DefaultValue = 0x0000FF000, HelpText = "Цвет пометки добавленных фрагментов")]
		public int ColorAdded { get; set; }
		[Option('d', "deleted", DefaultValue = 0x000000FF, HelpText = "Цвет пометки удалённых фрагментов")]
		public int ColorDeleted { get; set; }

		//--------------------------------------------------------------------------------------------------
		[HelpOption]
		public string GetUsage()
		{
			HelpText ht = HelpText.AutoBuild(this, (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));

			var sb = new StringBuilder();
			sb.AppendLine();
			sb.AppendLine("Применение:");
			sb.AppendLine("PDFCompare PDF1 PDF2 [PDFResult] [Параметры]");
			sb.AppendLine("	PDF1 PDF2 - документы для сравнения");
			sb.AppendLine("	PDFResult - отчёт по результатам сравнения. Если не указан, то автоматически формируется рядом с PDF2.");
			sb.AppendLine();

			sb.Append("Параметры:");

			ht.Heading = "PDFCompare. Сравнение pdf файлов.";
			ht.Copyright = sb.ToString(); // Replace copyright with commands info
			return ht;
		}
	}
}
