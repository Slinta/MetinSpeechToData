using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using Metin2SpeechToData;
using static SheetSync.Structures;
using static Metin2SpeechToData.Spreadsheet.SsConstants;
using System.Linq;
using System.Windows.Forms;

namespace SheetSync {


	class DiffChecker {
		public Typos[] typos;
		private readonly HotKeyMapper m = new HotKeyMapper();


		public DiffChecker(ExcelPackage sheetsFile, FileInfo[] currFiles) {
			Item[] allItems = GetItems(currFiles);
			getDiffs = FindDiffs(allItems.ToDictionary((str) => str.name, (data) => data), sheetsFile);

			Console.WriteLine();
			Console.WriteLine("Found " + getDiffs.Length + " differences:");

			foreach (Diffs diff in getDiffs) {
				Console.WriteLine("Item named '{3}' in sheet '{2}' is assigned as {0} Yang, but as {1} Yang in definition!",
					diff.sheetYangVal.ToString("C").Replace(",00 Kč", ""),
					diff.fileYangVal.ToString("C").Replace(",00 Kč", ""),
					diff.currentSheet.Name,
					diff.currentSheet.GetValue<string>(diff.location.Row, diff.location.Column - 1));
			}

			if (getDiffs.Length > 0) {
				Console.WriteLine();
				Console.WriteLine("(1) Sync data from definition files into the sheet.");
				Console.WriteLine("(2) Sync data from spreadsheet into definition files.");
				m.AssignToHotkey(Keys.D1, 1, Selected);
				m.AssignToHotkey(Keys.D2, 2, Selected);
				Console.ReadLine();
			}
			sheetsFile.Save();
		}

		public Typos[] getTypos {
			get { return typos; }
		}

		public Diffs[] getDiffs { get; }

		private void Selected(int selection) {
			if (selection == 1) {
				//Definition into Sheet
				foreach (Diffs difference in getDiffs) {
					difference.currentSheet.SetValue(difference.location.Address, difference.fileYangVal);
				}
			}
			else {
				string fileName = "";
				string[] currFile;
				foreach (Diffs diff in getDiffs) {
					fileName = diff.itemFile;
					currFile = File.ReadAllLines(fileName);

					currFile[diff.itemDef] = currFile[diff.itemDef].Replace(diff.fileYangVal.ToString(), diff.sheetYangVal.ToString());
					File.WriteAllLines(fileName, currFile);
				}
			}
			Console.WriteLine("Done");
			Console.WriteLine("Press Enter to save changes and exit...");
		}

		private static Item[] GetItems(FileInfo[] currFiles) {
			List<Item> items = new List<Item>();
			for (int i = 0; i < currFiles.Length; i++) {
				string[] fileContent = File.ReadAllLines(currFiles[i].FullName);
				int index = 0;
				while (fileContent[index].StartsWith("\t") || fileContent[index].Contains("{") ||
					fileContent[index].StartsWith("}") || string.IsNullOrWhiteSpace(fileContent[index])) {
					index++;
				}
				while (index < fileContent.Length) {
					if (string.IsNullOrWhiteSpace(fileContent[index]) || fileContent[index].StartsWith("#")) {
						index++;
						continue;
					}

					string[] split = fileContent[index].Split(',');
					Item item = new Item(split[0].Split('/')[0], currFiles[i].FullName, index, uint.Parse(split[1]));
					items.Add(item);
					index++;
				}
			}
			return items.ToArray();
		}

		private Diffs[] FindDiffs(IReadOnlyDictionary<string, Item> items, ExcelPackage sheetsFile) {
			List<Diffs> diffs = new List<Diffs>();
			ExcelWorksheets sheets = sheetsFile.Workbook.Worksheets;
			int sheetIndex = 2;
			if (sheets[1].Name != H_DEFAULT_SHEET_NAME) {
				Console.WriteLine("Missing fist sheet '" + H_DEFAULT_SHEET_NAME + "'... treating it as data sheet.");
				sheetIndex = 1;
			}
			List<Typos> currTypos = new List<Typos>();
			while (sheetIndex <= sheets.Count) {
				ExcelWorksheet sheet = sheets[sheetIndex];
				ExcelCellAddress currAddr = new ExcelCellAddress(3, 1);
				while (sheet.Cells[currAddr.Address].Value != null) {

					string sheetItemName = sheet.Cells[currAddr.Address].GetValue<string>();
					try {
						if (items[sheetItemName].name == sheetItemName) { /*Check wheter all items are spelt properly*/ }
					}
					catch {
						Console.WriteLine("Found invalid entry at '" + currAddr.Address + "' with name: '" + sheetItemName + "' ... skipping");
						int min = 20;
						string curr_min = "";
						foreach (string fileName in items.Keys) {
							int dist = WordSimilarity.Compute(fileName, sheetItemName);
							if (dist < min) {
								min = dist;
								curr_min = fileName;
								continue;
							}
							if (dist == min) {
								curr_min = curr_min + "' or '" + fileName;
							}
						}
						Console.WriteLine("Possibly '" + curr_min + "'?");
						Typos t = new Typos(sheetItemName, curr_min.Replace("' or '", "_").Trim('\'', ' ').Split('_'), currAddr, sheet);
						currTypos.Add(t);
						currAddr = SpreadsheetHelper.Advance(sheet, currAddr);
						if (currAddr == null) {
							break;
						}
						continue;
					}
					ExcelCellAddress yangs = new ExcelCellAddress(currAddr.Row, currAddr.Column + 1);

					uint localYangVal = sheet.Cells[yangs.Address].GetValue<uint>();
					if (localYangVal != items[sheetItemName].yangValue) {
						diffs.Add(new Diffs(sheet, yangs, items[sheetItemName].fileOrigin, items[sheetItemName].fileLine, localYangVal, items[sheetItemName].yangValue));
					}


					currAddr = SpreadsheetHelper.Advance(sheet, currAddr);
					if (currAddr == null) {
						break;
					}
				}

				sheetIndex++;
			}
			typos = currTypos.ToArray();
			return diffs.ToArray();
		}
	}
}
