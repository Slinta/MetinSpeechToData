using System;
using System.IO;

namespace Metin2SpeechToData {
	public class Configuration {

		public FileInfo xlsxFile { get; private set; }

		private const string SHEET_NAME = "Metin2 Drop Speadsheet.xlsx";

		public Configuration(string filePath) {
			if (!File.Exists(filePath)) {
				Console.WriteLine("You are missing a configuration file, this happens when you start the application for the first time," +
								  "or you had deleted it.");
				RecreateConfig();
			}
			else {
				ParseConfig(filePath);
			}
		}

		private void RecreateConfig() {
			using (StreamWriter sw = File.CreateText(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg")) {
				bool yes = GetBoolInput("Do you want the .xlsx file in the current directory ?");
				if (yes) {
					xlsxFile = new FileInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + SHEET_NAME);
					OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(xlsxFile);
				}
				else {
					throw new Exception("TODO");
				}
				sw.Write("PATH{\n\t" + xlsxFile + "\n}");
			}
		}

		private void ParseConfig(string filePath) {
			using (StreamReader sr = File.OpenText(filePath)) {
				while (!sr.EndOfStream) {
					string line = sr.ReadLine();
					if (line.Contains("{")) {
						string[] seg = line.Split('{');
						switch (seg[0]) {
							case "PATH": {
								line = sr.ReadLine();
								line = line.Trim(' ', '\t', '\n');
								xlsxFile = new FileInfo(line);
								break;
							}
						}
					}
				}
				Console.WriteLine("Parsed");
			}
		}



		public static bool GetBoolInput(string question) {
			Console.WriteLine(question);
			string line = Console.ReadLine();
			if (line == "yes" || line == "y") {
				return true;
			}
			return false;
		}
	}
}
