using OfficeOpenXml;
using System;
using System.IO;
using System.Windows.Forms;

namespace Metin2SpeechToData {
	public class Configuration {

		public enum SheetViewer {
			EXCEL,
			CALC
		}

		public FileInfo xlsxFile { get; private set; }

		public const string DEFAULT_FILE_NAME = "Metin2 Drop Spreadsheet.xlsx";
		public const string AREA_REGEXP = @"(Mob_)?\w+$";
		public const string CHESTS_REGEXP = @"\w+\ (C|c)hest[+-]?";
		public const string FILE_EXT = ".definition";

		private const uint DEFAULT_STACK_DEPHT = 5;
		private const uint DEFAULT_INTERNAL_MODIFICATION_COUNT = 1;
		private const float DEFAULT_SPEECH_ACCEPTANCE_THRESHOLD = 0.8f;
		private const SheetViewer DEFAULT_SHEET_VIEWER = SheetViewer.EXCEL;
		private const int DEFAULT_AVERAGE_COMPUTATION_TIMESTEP = 15;
		private static int parsedTimeStampAverage = DEFAULT_AVERAGE_COMPUTATION_TIMESTEP;

		public static string sessionDirectory { get; private set; }
		public static uint undoHistoryLength { get; private set; } = DEFAULT_STACK_DEPHT;
		public static uint sheetChangesBeforeSaving { get; private set; } = DEFAULT_INTERNAL_MODIFICATION_COUNT;
		public static float acceptanceThreshold { get; private set; } = DEFAULT_SPEECH_ACCEPTANCE_THRESHOLD;
		public static SheetViewer sheetViewer { get; private set; } = DEFAULT_SHEET_VIEWER;
		public static TimeSpan minutesAverageDropValueInterval { get => new TimeSpan(0, parsedTimeStampAverage, 0);}
		public static bool debug { get; private set; } = false;


		public Configuration(string filePath) {
			if (!File.Exists(filePath)) {
				Console.WriteLine("You are missing a configuration file, this happens when you start the application for the first time, " +
								  "or you had deleted it.");
				RecreateConfig();
			}
			else {
				ParseConfig(filePath);
			}
			ValidateDirectory();
		}

		private void ValidateDirectory() {
			string commonDir = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar;

			if (!Directory.Exists(commonDir + "Definitions")) {
				Directory.CreateDirectory(commonDir + "Definitions");
			}
			if (!Directory.Exists(commonDir + "Hotkeys")) {
				Directory.CreateDirectory(commonDir + "Hotkeys");
			}
			if (!Directory.Exists(commonDir + "Sessions")) {
				Directory.CreateDirectory(commonDir + "Sessions");
			}
			if (!Directory.Exists(commonDir + "Templates")) {
				Directory.CreateDirectory(commonDir + "Templates");
			}
			sessionDirectory = commonDir + "Sessions" + Path.DirectorySeparatorChar;
		}

		private void RecreateConfig() {
			using (StreamWriter sw = File.CreateText(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg")) {
				bool yes = GetBoolInput("Do you want the .xlsx file in the current directory ?\ny/n");
				if (yes) {
					xlsxFile = new FileInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + DEFAULT_FILE_NAME);
					ExcelPackage package = new ExcelPackage(xlsxFile);
					SpreadsheetTemplates templates = new SpreadsheetTemplates();
					ExcelWorksheets sheets = templates.LoadTemplates();
					package.Workbook.Worksheets.Add("Metin2 Drop Analyzer",sheets["Main"]);
					package.SaveAs(xlsxFile);
					templates.Dispose();
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
					OfficeOpenXml.ExcelPackage package = new ExcelPackage(xlsxFile);
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
				sw.Write("# The threshold for recognition, if the recognizer is less confident in what you said than this value, it will ignore the word/sentence | DEAFULT=0.8 {0,5--0,95}\n");
				sw.Write("SPEECH_ACCEPTANCE_THRESHOLD= " + DEFAULT_SPEECH_ACCEPTANCE_THRESHOLD + "\n");
				sw.WriteLine();
				sw.Write("# Specify which program do you use for opening .xlsx files, because LO Calc and MS Excel are not fully compatible | DEFAUL=EXCEL {'EXCEL','CALC'}\n");
				sw.Write("DEFAULT_SHEET_EDITOR= " + DEFAULT_SHEET_VIEWER + "\n");
				sw.WriteLine();
				sw.Write("# Specify the interval for \"Average drop value per ____\" in session sheet (as minutes!) | DEFAULT=15 {5--60}\n");
				sw.Write("DEFAULT_AVERAGE_COMPUTATION_TIMESTEP= 15");
				sw.WriteLine();
			}
		}

		private void ParseConfig(string filePath) {
			bool parseSuccess = false;
			using (StreamReader sr = File.OpenText(filePath)) {
				if (sr.ReadLine() == "debug") {
					debug = true;
				}

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
													  "You have to replace the path to it in 'config.cfg' in this apps folder, or delete the configuration,\n");
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
							if (undoHistoryLength > 1000 && undoHistoryLength < 1) {
								throw new CustomException("The value for UNDO_HISTORY_LENGTH in 'config.cfg' is out of bounds!");
							}
						}
						else if (split[0] == "WRITE_XSLX_AFTER_NO_OF_MODIFICATIONS") {
							sheetChangesBeforeSaving = uint.Parse(split[1].Trim('\n', ' ', '\t'));
							if (sheetChangesBeforeSaving > 100 && sheetChangesBeforeSaving < 1) {
								throw new CustomException("The value for WRITE_XSLX_AFTER_NO_OF_MODIFICATIONS in 'config.cfg' is out of bounds!");
							}
						}
						else if (split[0] == "SPEECH_ACCEPTANCE_THRESHOLD") {
							acceptanceThreshold = float.Parse(split[1].Trim('\n', ' ', '\t'));
							if (acceptanceThreshold > 0.95f && acceptanceThreshold < 0.5f) {
								throw new CustomException("The value for SPEECH_ACCEPTANCE_THRESHOLD in 'config.cfg' is out of bounds!");
							}
						}
						else if (split[0] == "DEFAULT_SHEET_EDITOR") {
							if (!Enum.TryParse(split[1].Trim('\n', ' ', '\t'), out SheetViewer _sheetViewer)) {
								throw new CustomException("The value for DEFAULT_SHEET_EDITOR in 'config.cfg' is not 'CALC' or 'EXCEL'!");
							}
							else {
								sheetViewer = _sheetViewer;
							}
						}
						else if (split[0] == "DEFAULT_AVERAGE_COMPUTATION_TIMESTEP") {
							parsedTimeStampAverage = int.Parse(split[1].Trim('\n', ' ', '\t'));
							if (parsedTimeStampAverage > 60 && parsedTimeStampAverage < 5) {
								throw new CustomException("The value for DEFAULT_AVERAGE_COMPUTATION_TIMESTEP in 'config.cfg' is out of bounds!");
							}
						}
					}
				}
				if (parseSuccess) {
					if (debug) {
						Console.WriteLine("Parsed successfuly!");
					}
					return;
				}
				if(sr.ReadToEnd() == "") {
					Console.WriteLine("Configuration file is empty, and had to be deleted, restart this application.\nPress 'Enter' to exit...");
					Console.ReadLine();
					sr.Close();
					File.Delete(filePath);
					sr.Dispose();
					Environment.Exit(0);
				}
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
