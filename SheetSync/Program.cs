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
				throw new FileNotFoundException("Unable to find config file");
			}

			DirectoryInfo currentDirectory = new DirectoryInfo(Directory.GetCurrentDirectory());
			if (currentDirectory.GetDirectories("Definitions").Length == 0) {
				throw new InvalidOperationException("Not inside the program's directory");
			}

			FileInfo[] currFiles = currentDirectory.GetDirectories("Definitions")[0].GetFiles("*.definition");
			ExcelPackage sheetsFile = new ExcelPackage(cfg.xlsxFile);
			FileInfo[] sessionFiles = currentDirectory.GetDirectories("Sessions")[0].GetFiles("*.xlsx");

			DiffChecker checker = new DiffChecker(sheetsFile, currFiles);
			TypoResolution typoRes = new TypoResolution(checker.getTypos);
			MergeHelper merge = new MergeHelper(sheetsFile, sessionFiles);


			if (checker.getDiffs.Length == 0 && typoRes.getTypos.Length == 0 && sessionFiles.Length == 0) {
				Console.WriteLine("Evertying looks the way it should ;]\nPress enter to quit");
				Console.ReadLine();
				Environment.Exit(0);
			}

			sheetsFile.Save();
			Console.WriteLine("All done");
		}
	}
}

