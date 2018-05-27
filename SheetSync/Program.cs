using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using Metin2SpeechToData;
using System.Windows.Forms;
using System.Threading;
using static Metin2SpeechToData.Spreadsheet.SsConstants;
using static SheetSync.Structures;

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

			sessions = currentDirectory.GetDirectories("Sessions")[0].GetFiles("*.xlsx");
			List<int> unmergedFiles = new List<int>();
			for (int i = 0; i < sessions.Length; i++) {
				if (sessions[i].Attributes == FileAttributes.Archive) {
					unmergedFiles.Add(i);
				}
			}

			Item[] allItems = GetItems(currFiles);
			differences = FindDiffs(allItems.ToDictionary((str) => str.name, (data) => data), sheetsFile);
			if (differences.Length == 0 && typos.Length == 0 && sessions.Length == 0) {
				Console.WriteLine("Evertying looks the way it should ;]\nPress enter to quit");
				Console.ReadLine();
				Environment.Exit(0);
			}

			Console.WriteLine();
			Console.WriteLine("Found " + differences.Length + " differences:");

			foreach (Diffs diff in differences) {
				Console.WriteLine("Item named '{3}' in sheet '{2}' is assigned as {0} Yang, but as {1} Yang in definition!",
					diff.sheetYangVal.ToString("C").Replace(",00 Kč", ""),
					diff.fileYangVal.ToString("C").Replace(",00 Kč", ""),
					diff.currentSheet.Name,
					diff.currentSheet.GetValue<string>(diff.location.Row, diff.location.Column - 1));
			}

			if (typos.Length > 0 && Confirmation.WrittenConfirmation("Resovle name typos?")) {
				ResolveTypos();
			}


			if (differences.Length > 0) {
				Console.WriteLine();
				Console.WriteLine("(1) Sync data from definition files into the sheet.");
				Console.WriteLine("(2) Sync data from spreadsheet into definition files.");
				m.AssignToHotkey(Keys.D1, 1, Selected);
				m.AssignToHotkey(Keys.D2, 2, Selected);
				Console.ReadLine();
			}
			sheetsFile.Save();

			if (unmergedFiles.Count > 0) {
				Console.WriteLine("Found unmerged session files in 'Sessions' folder...");
				ExcelPackage mainP = new ExcelPackage(new FileInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + Configuration.DEFAULT_FILE_NAME));
				ExcelWorksheet main = mainP.Workbook.Worksheets[H_DEFAULT_SHEET_NAME];
				foreach (int i in unmergedFiles) {
					if (Confirmation.WrittenConfirmation("Merge '" + sessions[i].Name + "' with main file?")) {
						//TODO
						//sessions[i].Attributes = FileAttributes.Normal;
					}
				}
			}
			sheetsFile.SaveAs(new FileInfo(sheetsFile.File.FullName.Replace(".xlsx", "_NEW.xlsx"))); //TODO once it works, overwrite
			Console.WriteLine("All done");
		}

		private static void Selected(int selection) {
			if (selection == 1) {
				//Definition into Sheet
				foreach (Diffs difference in differences) {
					difference.currentSheet.SetValue(difference.location.Address, difference.fileYangVal);
				}
			}
			else {
				string fileName = "";
				string[] currFile;
				foreach (Diffs diff in differences) {
					fileName = diff.itemFile;
					currFile = File.ReadAllLines(fileName);

					currFile[diff.itemDef] = currFile[diff.itemDef].Replace(diff.fileYangVal.ToString(), diff.sheetYangVal.ToString());
					File.WriteAllLines(fileName, currFile);
				}
			}
			Console.WriteLine("Done");
			Console.WriteLine("Press Enter to save changes and exit...");
		}
	}
}

