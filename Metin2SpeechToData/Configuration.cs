using System;
using System.IO;
using System.Windows.Forms;

namespace Metin2SpeechToData {
	public class Configuration {

		public FileInfo xlsxFile { get; private set; }

		public const string DEFAULT_FILE_NAME = "Metin2 Drop Speadsheet.xlsx";
		public const string AREA_REGEXP = @"(Mob_)?\w+$";
		public const string CHESTS_REGEXP = @"\w+\ (C|c)hest[+-]?";
		public const string FILE_EXT = ".definition";

		private const uint DEFAULT_STACK_DEPHT = 5;
		private const uint DEFAULT_INTERNAL_MODIFICATION_COUNT = 1;
		private const float DEFAULT_SPEECH_ACCEPTANCE_THRESHOLD = 0.8f;

		public static uint undoHistoryLength { get; private set; } =  DEFAULT_STACK_DEPHT;
		public static uint sheetChangesBeforeSaving { get; private set; } = DEFAULT_INTERNAL_MODIFICATION_COUNT;
		public static float acceptanceThreshold { get; private set; } = DEFAULT_SPEECH_ACCEPTANCE_THRESHOLD;

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
					xlsxFile = new FileInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + DEFAULT_FILE_NAME);
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
						File.Delete(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg");
						Console.ReadLine();
						Environment.Exit(1);
						return;
					}
					xlsxFile = new FileInfo(folderBrowser.SelectedPath + Path.DirectorySeparatorChar + DEFAULT_FILE_NAME);
					OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(xlsxFile);
					package.Workbook.Worksheets.Add("Metin2 Drop Analyzer");
					package.SaveAs(xlsxFile);
				}
				sw.Write("# Path to the file\n");
				sw.Write("PATH{\n\t" + xlsxFile + "\n}\n");
				sw.WriteLine();
				sw.Write("# Undo stack depth, the amout of \"undos\" you can make | DEAFULT=5 {1--1000}\n");
				sw.Write("UNDO_HISTORY_LENGTH= " + DEFAULT_STACK_DEPHT + "\n");
				sw.WriteLine();
				sw.Write("# How many changes are made internally before writing into the file, lower is better for stability reasons, higher is less CPU intesive | DEAFULT=1 {1--100}\n");
				sw.Write("WRITE_XSLX_AFTER_NO_OF_MODIFICATIONS= " + DEFAULT_INTERNAL_MODIFICATION_COUNT + "\n");
				sw.WriteLine();
				sw.Write("# The threshold for recognition, if the recognizer is less confident in waht you said than this value, it will ignore the word/sentence | DEAFULT=0.8 {0,5--0,95}\n");
				sw.Write("SPEECH_ACCEPTANCE_THRESHOLD= " + DEFAULT_SPEECH_ACCEPTANCE_THRESHOLD + "\n");
				sw.WriteLine();
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
							if(undoHistoryLength > 1000 && undoHistoryLength < 1) {
								throw new CustomException("THe value for UNDO_HISTORY_LENGTH in 'config.cfg' is out of bounds!");
							}
						}
						else if (split[0] == "WRITE_XSLX_AFTER_NO_OF_MODIFICATIONS") {
							sheetChangesBeforeSaving = uint.Parse(split[1].Trim('\n', ' ', '\t'));
							if (sheetChangesBeforeSaving > 100 && sheetChangesBeforeSaving < 1) {
								throw new CustomException("THe value for WRITE_XSLX_AFTER_NO_OF_MODIFICATIONS in 'config.cfg' is out of bounds!");
							}
						}
						else if (split[0] == "SPEECH_ACCEPTANCE_THRESHOLD") {
							acceptanceThreshold = float.Parse(split[1].Trim('\n', ' ', '\t'));
							if (acceptanceThreshold > 0.95f && acceptanceThreshold < 0.5f) {
								throw new CustomException("THe value for SPEECH_ACCEPTANCE_THRESHOLD in 'config.cfg' is out of bounds!");
							}
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
