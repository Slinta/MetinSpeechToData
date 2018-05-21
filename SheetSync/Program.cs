using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using Metin2SpeechToData;
using System.Windows.Forms;
using System.Threading;

namespace Sheet_DefinitionValueSync {
	internal static class Program {

		private static Diffs[] differences;
		private static Typos[] typos;
		private static FileInfo[] sessions;
		private static HotKeyMapper m = new HotKeyMapper();
		private static ManualResetEventSlim evnt = new ManualResetEventSlim();

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

			Item[] allItems = GetItems(currFiles);
			differences = FindDiffs(allItems.ToDictionary((str) => str.name, (data) => data), sheetsFile);
			if (differences.Length == 0 && typos.Length == 0 && sessions.Length == 0) {
				Console.WriteLine("Evertying looks the way it should ;]\nPress enter to quit");
				Console.ReadLine();
				Environment.Exit(0);
			}

			if(sessions.Length > 0) {
				Console.WriteLine("Found session files in 'Sessions' folder...");
				if(Confirmation.WrittenConfirmation("Merge with main file (create new sheets in main file with time stamp)?")) {
					Console.WriteLine("Sorry, not suppoted yet!");
					//TODO
				}
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
			Console.WriteLine("All done");
		}

		private static void ResolveTypos() {
			m.hotkeyOverriding = true;
			for (int j = 0; j < typos.Length; j++) {
				evnt.Reset();
				Console.WriteLine("Typo: " + typos[j].originalTypo);
				Console.Write("Alternatives: ");
				for (int i = 0; i < typos[j].alternatives.Length - 1; i++) {
					Console.Write("(" + (i + 1) + ")-" + typos[j].alternatives[i] + ", ");
					m.AssignToHotkey(Keys.D1 + i, i+ 1, Resolve);
				}
				Console.WriteLine("(" + (typos[j].alternatives.Length) + ")-" + typos[j].alternatives[typos[j].alternatives.Length - 1]);
				m.AssignToHotkey(Keys.D1 + typos[j].alternatives.Length - 1, typos[j].alternatives.Length - 1, Resolve);
				evnt.Wait();
			}
		}

		private static int currentIndex = 0;
		private static void Resolve(int selected) {
			Console.WriteLine("Replaced '" + typos[currentIndex].originalTypo+"' with '"+ typos[currentIndex].alternatives[selected] + "'");
			typos[currentIndex].sheet.SetValue(typos[currentIndex].location.Address, typos[currentIndex].alternatives[selected]);
			evnt.Set();
			currentIndex++;
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
			List<Typos> currTypos = new List<Typos>();
			while (sheetIndex <= sheets.Count) {
				ExcelWorksheet sheet = sheets[sheetIndex];
				ExcelCellAddress currAddr = new ExcelCellAddress(ST_ROW, 1);
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
						Typos t = new Typos(sheetItemName, curr_min.Replace("' or '", "_").Trim('\'', ' ').Split('_'),currAddr, sheet);
						currTypos.Add(t);
						currAddr = Advance(sheet, currAddr);
						if (currAddr == default(ExcelCellAddress)) {
							break;
						}
						continue;
					}
					ExcelCellAddress yangs = new ExcelCellAddress(currAddr.Row, currAddr.Column + 1);

					uint localYangVal = sheet.Cells[yangs.Address].GetValue<uint>();
					if (localYangVal != items[sheetItemName].yangValue) {
						diffs.Add(new Diffs(sheet, yangs, items[sheetItemName].fileOrigin, items[sheetItemName].fileLine, localYangVal, items[sheetItemName].yangValue));
					}


					currAddr = Advance(sheet, currAddr);
					if (currAddr == default(ExcelCellAddress)) {
						break;
					}
				}

				sheetIndex++;
			}
			typos = currTypos.ToArray();
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

	internal struct Typos {
		public Typos(string word, string[] possible_replacements, ExcelCellAddress sheetLocation, ExcelWorksheet sheet) {
			originalTypo = word;
			alternatives = possible_replacements;
			location = sheetLocation;
			this.sheet = sheet;
		}
		public string originalTypo { get; }
		public string[] alternatives { get; }
		public ExcelCellAddress location { get; }
		public ExcelWorksheet sheet{ get; }

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
		public Diffs(ExcelWorksheet sheet, ExcelCellAddress address, string fileOrigin, int fileLine, uint sheetVal, uint fileVal) {
			currentSheet = sheet;
			location = address;
			itemDef = fileLine;
			itemFile = fileOrigin;
			sheetYangVal = sheetVal;
			fileYangVal = fileVal;
		}

		public string itemFile { get; }
		public ExcelWorksheet currentSheet { get; }
		public ExcelCellAddress location { get; }
		public int itemDef { get; }
		public uint sheetYangVal { get; }
		public uint fileYangVal { get; }
	}
}

