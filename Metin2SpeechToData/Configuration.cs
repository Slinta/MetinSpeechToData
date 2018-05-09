using System;
using System.IO;
using System.Windows.Forms;

namespace Metin2SpeechToData {
	public class Configuration {

		public FileInfo xlsxFile { get; private set; }

		public const string FILE_NAME = "Metin2 Drop Speadsheet.xlsx";
		public const uint DEFAULT_STACK_DEPHT = 5;
		public const uint DEFAULT_INTERNAL_MODIFICATION_COUNT = 1;

		public static uint undoHistoryLength { get; private set; } =  DEFAULT_STACK_DEPHT;
		public static uint sheetChangesBeforeSaving { get; private set; } = DEFAULT_INTERNAL_MODIFICATION_COUNT;

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
				bool yes = GetBoolInput("Do you want the .xlsx file in the current directory ?\ny/n");
				if (yes) {
					xlsxFile = new FileInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + FILE_NAME);
					OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(xlsxFile);
					package.Workbook.Worksheets.Add("Metin2 Drop Analyzer");
					package.SaveAs(xlsxFile);
				}
				else {
					FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
					folderBrowser.Description = "Select the directory where you want to put your .xlsx";
					folderBrowser.ShowNewFolderButton = true;
					folderBrowser.ShowDialog();
					if (string.IsNullOrWhiteSpace(folderBrowser.SelectedPath)) {
						Console.WriteLine("No path selected, quitting...\nPress 'Enter' to close the window");
						sw.Close();
						sw.Dispose();
						File.Delete(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg");
						Console.ReadLine();
						Environment.Exit(1);
						return;
					}
					xlsxFile = new FileInfo(folderBrowser.SelectedPath + Path.DirectorySeparatorChar + FILE_NAME);
					OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(xlsxFile);
					package.Workbook.Worksheets.Add("Metin2 Drop Analyzer");
					package.SaveAs(xlsxFile);
				}
				sw.Write("# Path to the file\n");
				sw.Write("PATH{\n\t" + xlsxFile + "\n}\n");
				sw.WriteLine();
				sw.Write("# Undo stack depth, the amout of \"undos\" you can make | DEAFULT=5\n");
				sw.Write("UNDO_HISTORY_LENGTH= " + DEFAULT_STACK_DEPHT + "\n");
				sw.WriteLine();
				sw.Write("# How many changes are made internally before writing into the file, lower is better for stability reasons, higher is less CPU intesive | DEAFULT=1\n");
				sw.Write("WRITE_XSLX_AFTER_NO_OF_MODIFICATIONS= " + DEFAULT_INTERNAL_MODIFICATION_COUNT + "\n");
			}
		}

		private void ParseConfig(string filePath) {
			bool parseSuccess = false;
			using (StreamReader sr = File.OpenText(filePath)) {
				while (!sr.EndOfStream) {
					string line = sr.ReadLine();
					if (string.IsNullOrWhiteSpace(line) || line.Contains("#")) {
						continue;
					}
					if (line.Contains("{")) {
						string[] split = line.Split('{');
						switch (split[0]) {
							case "PATH": {
								line = sr.ReadLine();
								line = line.Trim(' ', '\t', '\n');
								if (!File.Exists(line)) {
									Console.WriteLine("The file " + line + " was not found!\n" +
													  "You have to raplace the path to it in 'config.cfg' in this apps folder, or delete the configuration,\n" +
													  "new one, along with the sheet will be generated on restart.");
									Console.ReadKey();
									Environment.Exit(0);
									return;
								}
								xlsxFile = new FileInfo(line);
								parseSuccess = true;
								break;
							}
						}
					}
					if (line.Contains("=")) {
						string[] split = line.Split('=');
						if (split[0] == "UNDO_HISTORY_LENGTH") {
							undoHistoryLength = uint.Parse(split[1].Trim('\n', ' ', '\t'));
						}
						else if (split[0] == "WRITE_XSLX_AFTER_NO_OF_MODIFICATIONS") {
							sheetChangesBeforeSaving = uint.Parse(split[1].Trim('\n', ' ', '\t'));
						}
					}
				}
				if (parseSuccess) {
					if (Program.debug) {
						Console.WriteLine("Parsed successfuly!");
					}
					return;
				}
				throw new CustomException("Corrupted 'config.cfg' found in application directory. Delete it and restart!");
			}
		}

		/// <summary>
		/// Asks user a true/false 'question' returns true on 'yes' or 'y' else false
		/// </summary>
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
