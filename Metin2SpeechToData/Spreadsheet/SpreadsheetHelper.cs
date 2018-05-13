using System.Collections.Generic;
using OfficeOpenXml;
using Metin2SpeechToData.Structures;


namespace Metin2SpeechToData {
	public class SpreadsheetHelper {

		public const string DEFAULT_SHEET = "Metin2 Drop Analyzer";
		private readonly SpreadsheetInteraction main;


		public SpreadsheetHelper(SpreadsheetInteraction interaction) {
			main = interaction;
		}

		/// <summary>
		/// Adjusts collumn width of current sheet
		/// </summary>
		public void AutoAdjustColumns(Dictionary<string, SpreadsheetInteraction.Group>.ValueCollection values) {
			double currMaxWidth = 0;
			foreach (SpreadsheetInteraction.Group g in values) {
				for (int i = 0; i < g.totalEntries; i++) {
					main.currentSheet.Cells[g.elementNameFirstIndex.Row + i, g.elementNameFirstIndex.Column].AutoFitColumns();
					if (main.currentSheet.Column(g.elementNameFirstIndex.Column).Width >= currMaxWidth) {
						currMaxWidth = main.currentSheet.Column(g.elementNameFirstIndex.Column).Width;
					}
				}
				main.currentSheet.Column(g.elementNameFirstIndex.Column).Width = currMaxWidth;
				currMaxWidth = 0;

				for (int i = 0; i < g.totalEntries; i++) {
					int s = main.currentSheet.GetValue<int>(g.yangValueFirstIndex.Row + i, g.yangValueFirstIndex.Column);
					main.currentSheet.Column(g.yangValueFirstIndex.Column).Width = GetCellWidth(s, false);
					if (main.currentSheet.Column(g.yangValueFirstIndex.Column).Width > currMaxWidth) {
						currMaxWidth = main.currentSheet.Column(g.yangValueFirstIndex.Column).Width;
					}
				}
				main.currentSheet.Column(g.yangValueFirstIndex.Column).Width = currMaxWidth;
				currMaxWidth = 0;
			}
			main.Save();
		}


		/// <summary>
		/// Parses spreadsheet
		/// </summary>
		/// <param name="book">Current WorkBook</param>
		/// <param name="sheetName">Name of the sheet as it appears in Excel</param>
		/// <param name="type">Spreadsheet type to determine parsing method</param>
		public static Dicts LoadSpreadsheet(ExcelWorkbook book, string sheetName, SpreadsheetTemplates.SpreadsheetPresetType type) {
			ExcelWorksheet sheet = book.Worksheets[sheetName];
			switch (type) {
				case SpreadsheetTemplates.SpreadsheetPresetType.MAIN: {
					return default(Dicts);
				}
				case SpreadsheetTemplates.SpreadsheetPresetType.AREA: {
					Dicts d = new Dicts(true);
					DefinitionParserData data = DefinitionParser.instance.currentGrammarFile;
					int[] rowOfEachGroup = new int[data.groups.Length];
					int[] columnOfEachGroup = new int[data.groups.Length];
					int columnOffset = 4;
					int groupcounter = 0;
					foreach (string group in data.groups) {
						rowOfEachGroup[groupcounter] = 2;
						columnOfEachGroup[groupcounter] = groupcounter * columnOffset + 1;
						ExcelCellAddress address = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter]);
						groupcounter += 1;
						SpreadsheetInteraction.Group g = new SpreadsheetInteraction.Group(address, new ExcelCellAddress(address.Row + 1, address.Column),
																						  new ExcelCellAddress(address.Row + 1, address.Column + 1),
																						  new ExcelCellAddress(address.Row + 1, address.Column + 2));
						d.groups.Add(group, g);
					}
					foreach (DefinitionParserData.Item entry in data.entries) {
						SpreadsheetInteraction.Group g = d.groups[entry.group];
						g.totalEntries++;
						d.groups[entry.group] = g;
						for (int i = 0; i < data.groups.Length; i++) {
							if (data.groups[i] == entry.group) {
								groupcounter = i;
							}
						}
						rowOfEachGroup[groupcounter] += 1;
						ExcelCellAddress collected = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter] + 2);

						d.addresses.Add(entry.mainPronounciation, collected);
					}
					return d;
				}
				case SpreadsheetTemplates.SpreadsheetPresetType.ENEMY: {
					Dicts d = new Dicts(true);
					// First item entry at A2
					ExcelCellAddress baseAddr = new ExcelCellAddress("A2");
					ExcelCellAddress current = baseAddr;

					bool EOF = false;
					while (!EOF) {
						if (sheet.Cells[current.Row, current.Column].Value == null) {
							current = new ExcelCellAddress(2, current.Column + 4);
							if (sheet.Cells[current.Address].Value == null) {
								EOF = true;
								continue;
							}
						}
						d.addresses.Add(sheet.Cells[current.Row, current.Column].GetValue<string>(), new ExcelCellAddress(current.Row, current.Column + 2));
						current = new ExcelCellAddress(current.Row + 1, current.Column);
					}
					return d;
				}
				default: {
					throw new CustomException("Uncathegorized sheet entered!");
				}
			}
		}

		/// <summary>
		/// Cathegorizes sheet based on in which file its name is located
		/// </summary>
		public SpreadsheetTemplates.SpreadsheetPresetType GetSheetType(string sheetName) {
			foreach (string locationName in DefinitionParser.instance.getDefinitionNames) {
				if (locationName == sheetName) {
					return SpreadsheetTemplates.SpreadsheetPresetType.AREA;
				}
			}

			foreach (MobParserData.Enemy enemy in DefinitionParser.instance.currentMobGrammarFile.enemies) {
				if (enemy.mobMainPronounciation == sheetName) {
					return SpreadsheetTemplates.SpreadsheetPresetType.ENEMY;
				}
			}
			return SpreadsheetTemplates.SpreadsheetPresetType.MAIN;
		}

		private double GetCellWidth(int number, bool addCurrencyOffset) {
			int count = DigitCount(number);
			int spaces = count / 3;
			double width = addCurrencyOffset ? 4 : 2;
			width += spaces + count;
			return width;
		}

		private int DigitCount(int i) {
			return i.ToString().Length;
		}

		#region Formula functions
		/// <summary>
		/// Sums cells in 'range' and puts result into 'result'
		/// </summary>
		public void Sum(ExcelCellAddress result, ExcelRange range) {
			main.currentSheet.Cells[result.Address].Formula = "SUM(" + range.Start.Address + ':' + range.End.Address + ")";
		}

		/// <summary>
		/// Sums cells in scattered range 'addresses' and puts result into 'result'
		/// </summary>
		public void Sum(ExcelCellAddress result, ExcelCellAddress[] addresses) {
			string s = "";
			foreach (ExcelCellAddress addr in addresses) {
				s = string.Join(",", s, addr.Address);
			}
			s = s.TrimStart(',');
			main.currentSheet.Cells[result.Address].Formula = "SUM(" + s + ")";
		}

		/// <summary>
		/// Averages cells in 'range' and puts result into 'result'
		/// </summary>
		public void Average(ExcelCellAddress result, ExcelRange range) {
			main.currentSheet.Cells[result.Address].Formula = "AVERAGE(" + range + ")";
		}

		/// <summary>
		/// Averages cells in scattered range 'addresses' and puts result into 'result'
		/// </summary>
		public void Average(ExcelCellAddress result, ExcelCellAddress[] addresses) {
			string s = "";
			foreach (ExcelCellAddress addr in addresses) {
				s = string.Join(",", s, addr.Address);
			}
			s = s.TrimStart(',');
			main.currentSheet.Cells[result.Address].Formula = "AVERAGE(" + s + ")";
		}

		/// <summary>
		/// Preforms a Sum of 'numerator' and divides it by a value in 'denominator' puts the result into 'result'
		/// </summary>
		public void DivideBy(ExcelCellAddress result, ExcelRange numerator, ExcelCellAddress denominator) {
			main.currentSheet.Cells[result.Address].Formula = "SUM(" + numerator.Start.Address + ':' + numerator.End.Address + ")/" + denominator.Address;
		}

		/// <summary>
		/// Gets continuous range between two addressess
		/// </summary>
		public ExcelRange GetRangeContinuous(ExcelCellAddress addr1, ExcelCellAddress addr2) {
			main.currentSheet.Select(addr1.Address + ":" + addr2.Address);
			ExcelRange range = main.currentSheet.SelectedRange;
			main.currentSheet.Select("A1");
			return range;
		}

		/// <summary>
		/// Gets continuous range between two addressess as strings
		/// </summary>
		public ExcelRange GetRangeContinuous(string addr1, string addr2) {
			main.currentSheet.Select(addr1 + ":" + addr2);
			ExcelRange range = main.currentSheet.SelectedRange;
			main.currentSheet.Select("A1");
			return range;
		}
		#endregion
	}
}
