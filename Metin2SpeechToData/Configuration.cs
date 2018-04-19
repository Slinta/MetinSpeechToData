﻿using System;
using System.IO;
using System.Windows.Forms;

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
				bool yes = GetBoolInput("Do you want the .xlsx file in the current directory ?\ny/n");
				if (yes) {
					xlsxFile = new FileInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + SHEET_NAME);
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
						folderBrowser.ShowDialog();
					}
					if (string.IsNullOrWhiteSpace(folderBrowser.SelectedPath)) {
						Environment.Exit(0);
					}
					xlsxFile = new FileInfo(folderBrowser.SelectedPath + Path.DirectorySeparatorChar + SHEET_NAME);
					OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(xlsxFile);
					package.Workbook.Worksheets.Add("Metin2 Drop Analyzer");
					package.SaveAs(xlsxFile);
				}
				sw.Write("PATH{\n\t" + xlsxFile + "\n}");
			}
		}

		private void ParseConfig(string filePath) {
			bool parseSuccess = false;
			using (StreamReader sr = File.OpenText(filePath)) {
				while (!sr.EndOfStream) {
					string line = sr.ReadLine();
					if (line.Contains("{")) {
						string[] seg = line.Split('{');
						switch (seg[0]) {
							case "PATH": {
								line = sr.ReadLine();
								line = line.Trim(' ', '\t', '\n');
								if (!File.Exists(line)) {
									Console.WriteLine("The file " + line + " was not found!\n" +
													  "You have to raplace the path to it in 'config.cfg' or delete the config,\n" +
													  "and new cfg + sheet will be generated on restart.");
									Console.ReadKey();
									Environment.Exit(0);
									break;
								}
								xlsxFile = new FileInfo(line);
								parseSuccess = true;
								break;
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
