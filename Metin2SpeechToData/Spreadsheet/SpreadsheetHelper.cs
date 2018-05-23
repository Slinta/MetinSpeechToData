using System.Collections.Generic;
using OfficeOpenXml;
using Metin2SpeechToData.Structures;
using static Metin2SpeechToData.Spreadsheet.SsConstants;
using System;

namespace Metin2SpeechToData {
	public class SpreadsheetHelper {

		private readonly SpreadsheetInteraction interaction;

		public SpreadsheetHelper(SpreadsheetInteraction interaction) {
			this.interaction = interaction;
		}

		/// <summary>
		/// Adjusts collumn width of current sheet
		/// </summary>
		public void AutoAdjustColumns(Dictionary<string, SpreadsheetInteraction.Group>.ValueCollection values) {
			double currMaxWidth = 0;
			foreach (SpreadsheetInteraction.Group g in values) {
				string groupStartAddress = g.elementNameFirstIndex.Address;

				for (int i = 0; i < g.totalEntries; i++) {
					interaction.currentSheet.Cells[OffsetAddress(groupStartAddress,i,0)].AutoFitColumns();
					if (interaction.currentSheet.Column(g.elementNameFirstIndex.Column).Width >= currMaxWidth) {
						currMaxWidth = interaction.currentSheet.Column(g.elementNameFirstIndex.Column).Width;
					}
				}

				interaction.currentSheet.Column(g.elementNameFirstIndex.Column).Width = currMaxWidth;
				currMaxWidth = 0;

				for (int i = 0; i < g.totalEntries; i++) {
					int s = interaction.currentSheet.GetValue<int>(g.elementNameFirstIndex.Row + i, g.elementNameFirstIndex.Column + 1);
					interaction.currentSheet.Column(g.elementNameFirstIndex.Column + 1).Width = GetCellWidth(s);
					if (interaction.currentSheet.Column(g.elementNameFirstIndex.Column + 1).Width > currMaxWidth) {
						currMaxWidth = interaction.currentSheet.Column(g.elementNameFirstIndex.Column + 1).Width;
					}
				}
				interaction.currentSheet.Column(g.elementNameFirstIndex.Column + 1).Width = currMaxWidth;
				currMaxWidth = 0;
			}
			interaction.Save();
		}

		/// <summary>
		/// If given address containing itemName
		/// this function will keep returning following addresses,
		/// until it reaches an empty cell in next data column, then it returns null
		/// <para>A3 > A4 > A5 .. Ax is null continue next column E3 > E4 ... I3 is empty >> NULL</para>
		/// </summary>
		public static ExcelCellAddress Advance(ExcelWorksheet sheet, ExcelCellAddress currAddr) {
			currAddr = new ExcelCellAddress(currAddr.Row + 1, currAddr.Column);
			if (sheet.Cells[currAddr.Address].Value == null) {
				currAddr = new ExcelCellAddress(H_FIRST_ROW, currAddr.Column + H_COLUMN_INCREMENT);
				if (sheet.Cells[currAddr.Address].Value == null) {
					return null;
				}
			}
			return currAddr;
		}

		/// <summary>
		/// Creates a hyperlink in 'currentSheet' at 'currentCellAddress' pointing to 'otherSheet' to 'locationAddress', hide link syntax with 'displeyText'
		/// </summary>
		public static void HyperlinkCell(ExcelWorksheet currentSheet, string currentCellAddress, ExcelWorksheet otherSheet, string locationAddress, string displayText) {
			if (Configuration.sheetViewer == Configuration.SheetViewer.EXCEL) {
				currentSheet.Cells[currentCellAddress].Formula =
					string.Format("HYPERLINK(\"#'{0}'!{1}\",\"{2}\")", otherSheet.Name, locationAddress, displayText);
			}
			else {
				currentSheet.Cells[currentCellAddress].Formula =
					string.Format("HYPERLINK(\"#{0}.{1}\",\"{2}\")", otherSheet.Name, locationAddress, displayText);
			}
			currentSheet.Cells[currentCellAddress].Calculate();
		}

		/// <summary>
		/// Parses spreadsheet
		/// </summary>
		/// <param name="book">Current WorkBook</param>
		/// <param name="sheetName">Name of the sheet as it appears in Excel</param>
		/// <param name="type">Spreadsheet type to determine parsing method</param>
		public static Dicts LoadSpreadsheet(ExcelWorkbook book, string sheetName, SpreadsheetTemplates.SpreadsheetPresetType type) {
			ExcelWorksheet sheet = book.Worksheets[sheetName];
			Dicts d = new Dicts(true);

			switch (type) {
				case SpreadsheetTemplates.SpreadsheetPresetType.AREA: {
					DefinitionParserData data = DefinitionParser.instance.currentGrammarFile;
					byte[] rowOfEachGroup = new byte[data.groups.Length];
					byte[] columnOfEachGroup = new byte[data.groups.Length];
					byte groupcounter = 0;
					foreach (string group in data.groups) {
						rowOfEachGroup[groupcounter] = 2;
						columnOfEachGroup[groupcounter] = (byte)(groupcounter * H_COLUMN_INCREMENT + 1);
						ExcelCellAddress address = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter]);
						groupcounter += 1;
						SpreadsheetInteraction.Group g = new SpreadsheetInteraction.Group(address, new ExcelCellAddress(address.Row + 1, address.Column));
						d.groups.Add(group, g);
					}
					foreach (DefinitionParserData.Item entry in data.entries) {
						SpreadsheetInteraction.Group g = d.groups[entry.group];
						g.totalEntries++;
						d.groups[entry.group] = g;
						for (byte i = 0; i < data.groups.Length; i++) {
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
					ExcelCellAddress current = new ExcelCellAddress("A2");

					while (current != null) {
						d.addresses.Add(sheet.Cells[current.Row, current.Column].GetValue<string>(), new ExcelCellAddress(current.Row, current.Column + 2));
						current = Advance(sheet, current);
					}
					return d;
				}
				default: {
					throw new CustomException("Uncathegorized sheet entered!");
				}
			}
		}

		/// <summary>
		/// Takes 'current' address and offsets it by 'rowOffset' rows and'colOffset' columns 
		/// </summary>
		public static string OffsetAddress(string current, int rowOffset, int colOffset) {
			ExcelCellAddress a = new ExcelCellAddress(current);
			return new ExcelCellAddress(a.Row + rowOffset, a.Column + colOffset).Address;
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

		/// <summary>
		/// Helper function for AutoAjdustComluns
		/// </summary>
		private double GetCellWidth(int number) {
			int count = number.ToString().Length;
			int spaces = count / 3;
			return spaces + count;
		}

		#region Formula functions
		/// <summary>
		/// Sums cells in 'range' and puts result into 'result'
		/// </summary>
		public void Sum(ExcelCellAddress result, ExcelRange range) {
			interaction.currentSheet.Cells[result.Address].Formula = "SUM(" + range.Start.Address + ':' + range.End.Address + ")";
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
			interaction.currentSheet.Cells[result.Address].Formula = "SUM(" + s + ")";
		}

		/// <summary>
		/// Averages cells in 'range' and puts result into 'result'
		/// </summary>
		public void Average(ExcelCellAddress result, ExcelRange range) {
			interaction.currentSheet.Cells[result.Address].Formula = "AVERAGE(" + range + ")";
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
			interaction.currentSheet.Cells[result.Address].Formula = "AVERAGE(" + s + ")";
		}

		/// <summary>
		/// Preforms a Sum of 'numerator' and divides it by a value in 'denominator' puts the result into 'result'
		/// </summary>
		public void DivideBy(ExcelCellAddress result, ExcelRange numerator, ExcelCellAddress denominator) {
			interaction.currentSheet.Cells[result.Address].Formula = "SUM(" + numerator.Start.Address + ':' + numerator.End.Address + ")/" + denominator.Address;
		}

		/// <summary>
		/// Gets continuous range between two addressess
		/// </summary>
		public ExcelRange GetRangeContinuous(ExcelCellAddress addr1, ExcelCellAddress addr2) {
			interaction.currentSheet.Select(addr1.Address + ":" + addr2.Address);
			ExcelRange range = interaction.currentSheet.SelectedRange;
			interaction.currentSheet.Select("A1");
			return range;
		}

		/// <summary>
		/// Gets continuous range between two addressess as strings
		/// </summary>
		public ExcelRange GetRangeContinuous(string addr1, string addr2) {
			interaction.currentSheet.Select(addr1 + ":" + addr2);
			ExcelRange range = interaction.currentSheet.SelectedRange;
			interaction.currentSheet.Select("A1");
			return range;
		}
		#endregion
	}
}
