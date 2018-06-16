using System;
using System.IO;
using OfficeOpenXml;
using Metin2SpeechToData;

namespace SheetSync {
	internal static class Program {

		[STAThread]
		private static void Main(string[] args) {
			Configuration cfg = new Configuration(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg");
			if (cfg.xlsxFile == null) {
				throw new FileNotFoundException("Unable to find config file!");
			}

			DirectoryInfo currentDirectory = new DirectoryInfo(Directory.GetCurrentDirectory());
			if (currentDirectory.GetDirectories("Definitions").Length == 0) {
				throw new InvalidOperationException("Not inside the program's directory!");
			}
			ExcelPackage sheetsFile = new ExcelPackage(cfg.xlsxFile);
			FileInfo[] sessionFiles = currentDirectory.GetDirectories("Sessions")[0].GetFiles("*.xlsx");

			ValidateMain validate = new ValidateMain(sheetsFile);
			MergeHelper merge = new MergeHelper(sheetsFile, sessionFiles);
			sheetsFile.Save();

			Console.WriteLine("Evertying looks the way it should ;]\nPress 'Enter' to exit...");
			Console.ReadLine();
		}
	}
}

