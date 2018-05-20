using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using Metin2SpeechToData;

namespace Sheet_DefinitionValueSync {
	internal static class Program {
		static void Main(string[] args) {
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

			Item[] allItems = GetItems(currFiles);
			Diffs[] differences = FindDiffs(allItems.ToDictionary((str) => str.name, (data) => data), sheetsFile);
			Console.WriteLine("Found " + differences.Length + " differences");
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

		private const int ST_ROW = 3;
		private const int COL_INC = 4;

		private static Diffs[] FindDiffs(IReadOnlyDictionary<string, Item> items, ExcelPackage sheetsFile) {
			List<Diffs> diffs = new List<Diffs>();
			ExcelWorksheets sheets = sheetsFile.Workbook.Worksheets;
			int sheetIndex = 2;
			if (sheets[1].Name != SpreadsheetHelper.DEFAULT_SHEET) {
				Console.WriteLine("Missing fist sheet '" + SpreadsheetHelper.DEFAULT_SHEET + "'... treating it as data sheet.");
				sheetIndex = 1;
			}

			while (sheetIndex <= sheets.Count) {
				ExcelWorksheet sheet = sheets[sheetIndex];
				ExcelCellAddress currAddr = new ExcelCellAddress(ST_ROW, 1);
				while (sheet.Cells[currAddr.Address].Value != null) {

					string name = sheet.Cells[currAddr.Address].GetValue<string>();
					try {
						if (items[name].name == name) { /*Check wheter all items are spelt properly*/ }
					}
					catch {
						Console.WriteLine("Found invalid entry at '" +currAddr.Address + "' with name: '" + name + "' ... skipping");
						currAddr = Advance(sheet, currAddr);
						if (currAddr == default(ExcelCellAddress)) {
							break;
						}
						continue;
					}
					ExcelCellAddress yangs = new ExcelCellAddress(currAddr.Row, currAddr.Column + 1);

					uint localYangVal = sheet.Cells[yangs.Address].GetValue<uint>();
					if (localYangVal != items[name].yangValue) {
						diffs.Add(new Diffs(sheet, yangs, items[name].fileOrigin, items[name].fileLine));
					}


					currAddr = Advance(sheet, currAddr);
					if(currAddr == default(ExcelCellAddress)) {
						break;
					}
				}

				sheetIndex++;
			}
			return diffs.ToArray();
		}

		private static ExcelCellAddress Advance(ExcelWorksheet sheet, ExcelCellAddress currAddr) {
			currAddr = new ExcelCellAddress(currAddr.Row + 1, currAddr.Column);
			if (sheet.Cells[currAddr.Address].Value == null) {
				currAddr = new ExcelCellAddress(ST_ROW, currAddr.Column + COL_INC);
				if (sheet.Cells[currAddr.Address].Value == null) {
					return default(ExcelCellAddress);
				}
			}
			return currAddr;
		}
	}

	internal struct Item {
		public Item(string name, string fileOrigin, int fileLine, uint yangValue) {
			this.name = name;
			this.fileOrigin = fileOrigin;
			this.fileLine = fileLine;
			this.yangValue = yangValue;
		}

		public string name { get; }
		public string fileOrigin { get; }
		public int fileLine { get; }
		public uint yangValue { get; }
	}

	internal struct Diffs {
		public Diffs(ExcelWorksheet sheet, ExcelCellAddress address, string fileOrigin, int fileLine) {
			currentSheet = sheet;
			location = address;
			itemDef = fileLine;
			itemFile = fileOrigin;
		}

		public string itemFile { get; }
		public ExcelWorksheet currentSheet { get; }
		public ExcelCellAddress location { get; }
		public int itemDef { get; }
	}
}

