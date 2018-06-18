using Metin2SpeechToData;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

using static Metin2SpeechToData.Configuration;
using static Metin2SpeechToData.Spreadsheet.SsConstants;


namespace SheetSync {
	class ValidateMain {

		private ExcelWorksheet mainSheet;
		private readonly ExcelWorkbook content;

		public ValidateMain(ExcelPackage main) {
			content = main.Workbook;
			mainSheet = content.Worksheets[1];
			System.Console.WriteLine("Validating main...");

			RecreateMain(main);
			Validate();

			System.Console.WriteLine("File Validated successfully");
		}

		/// <summary>
		/// Deletes first sheet in 'main' and replaces it with clean one
		/// </summary>
		public void RecreateMain(ExcelPackage main) {
			SpreadsheetTemplates t = new SpreadsheetTemplates();
			ExcelWorksheets sheets = t.LoadTemplates();
			content.Worksheets.Delete(1);
			content.Worksheets.Add(H_DEFAULT_SHEET_NAME, sheets["Main"]);
			content.Worksheets.MoveToStart(H_DEFAULT_SHEET_NAME);
			mainSheet = main.Workbook.Worksheets[1];
		}

		/// <summary>
		/// Goes through all sessions in Sessions folder and sorts them into cathegories in the first sheet
		/// </summary>
		public void Validate() {
			FileInfo[] sessions = new DirectoryInfo(sessionDirectory).GetFiles("*.xlsx");
			sessions.Sort(0, sessions.Length - 1, FileComparer);

			FileInfo[] files = new DirectoryInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Definitions").GetFiles("Mob_*.definition");

			MobParserData[] data = new MobParserData().Parse(files);
			ExcelCellAddress sheets = new ExcelCellAddress(MAIN_SHEET_LINKS);
			ExcelCellAddress mergedSessions = new ExcelCellAddress(MAIN_MERGED_LINKS);
			ExcelCellAddress unmergedSessions = new ExcelCellAddress(MAIN_UNMERGED_LINKS);

			ExcelWorksheet current = mainSheet;
			int sheetCount = content.Worksheets.Count;

			for (int i = 2; i <= sheetCount; i++) {
				bool found = false;
				string currName = content.Worksheets[i].Name;
				foreach (MobParserData mobData in data) {
					foreach (MobParserData.Enemy enemy in mobData.enemies) {
						if (enemy.mobMainPronounciation == currName) {
							found = true;
						}
					}
					if (found) {
						break;
					}
				}
				if (found) {
					continue;
				}

				while (current.GetValue(sheets.Row, sheets.Column) != null) {
					if (mainSheet.Cells[sheets.Address].Value.ToString() == currName) {
						found = true;
						break;
					}
					sheets = SpreadsheetHelper.OffsetAddress(sheets, 1, 0);
				}
				if (!found) {
					SpreadsheetHelper.Copy(mainSheet, sheets.Address, SpreadsheetHelper.OffsetAddress(sheets, 0, 3).Address,
													  SpreadsheetHelper.OffsetAddress(sheets, 1, 0).Address, SpreadsheetHelper.OffsetAddress(sheets, 1, 3).Address);
					SpreadsheetHelper.HyperlinkCell(mainSheet, sheets.Address, current, "A1", currName);
				}
			}

			foreach (FileInfo file in sessions) {
				string sessionName = SpreadsheetHelper.GetSessionName(file);
				if (file.Attributes != FileAttributes.Archive) {
					SortIntoCathegory(file, mergedSessions, sessionName);
				}
				else {
					SortIntoCathegory(file, unmergedSessions, sessionName);
				}
			}
		}

		private void SortIntoCathegory(FileInfo currentFile, ExcelCellAddress sessionAddr, string sessionName) {
			bool found = false;
			while (mainSheet.GetValue(sessionAddr.Row, sessionAddr.Column) != null) {
				if (mainSheet.Cells[sessionAddr.Address].Value.ToString() == sessionName) {
					found = true;
					break;
				}
				sessionAddr = SpreadsheetHelper.OffsetAddress(sessionAddr, 1, 0);
			}
			if (!found) {
				SpreadsheetHelper.Cut(mainSheet, sessionAddr.Address, SpreadsheetHelper.OffsetAddress(sessionAddr, 0, 3).Address,
						  SpreadsheetHelper.OffsetAddress(sessionAddr, 1, 0).Address, SpreadsheetHelper.OffsetAddress(sessionAddr, 1, 3).Address);
				SpreadsheetHelper.HyperlinkAcrossFiles(currentFile, "Session", "A1", mainSheet, sessionAddr.Address, sessionName);
				sessionAddr = SpreadsheetHelper.OffsetAddress(sessionAddr, 1, 0);
			}
		}

		private int FileComparer(FileInfo a, FileInfo b) {
			string nameA = SpreadsheetHelper.GetSessionName(a);
			string nameB = SpreadsheetHelper.GetSessionName(b);
			return string.Compare(nameA, nameB);
		}
	}
}
