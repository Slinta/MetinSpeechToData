using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public class SpreadsheetInteraction {

		private ExcelPackage xlsxFile;
		private ExcelWorkbook content;

		public ExcelWorksheet currentSheet { get; private set; }
		private Dictionary<string, ExcelCellAddress> nameLookupDictionary;

		/// <summary>
		/// Create new file and add initial sheet, set content and sheet file is saved in the process!
		/// </summary>
		public SpreadsheetInteraction(string path, string worksheetName) {
			xlsxFile = new ExcelPackage(new FileInfo(path));
			content = xlsxFile.Workbook;
			OpenWorksheet(worksheetName);
			xlsxFile.Save();
		}
		/// <summary>
		/// Create new file at 'path' and set constent, sheet it null!
		/// </summary>
		public SpreadsheetInteraction(string path) {
			xlsxFile = new ExcelPackage(new FileInfo(path));
			content = xlsxFile.Workbook;
		}

		/// <summary>
		/// Open existing worksheet or add one if it does't exist
		/// </summary>
		public void OpenWorksheet(string sheetname) {
			if (content.Worksheets[sheetname] == null) {
				content.Worksheets.Add(sheetname);
			}
			currentSheet = content.Worksheets[sheetname];
		}

		/// <summary>
		/// Add a sheet by index
		/// </summary>
		public void OpenWorksheet(int sheetindex) {
			if (content.Worksheets[sheetindex] == null) {
				throw new Exception("No worksheet with id " + sheetindex + " exists");
			}
			else {
				currentSheet = content.Worksheets[sheetindex];
			}
		}

		/// <summary>
		/// Adds a 'number' to the number cell at 'address' and saves the document
		/// </summary>
		public void AddNumberTo(ExcelCellAddress address, int number) {
			if (currentSheet == null) {
				throw new Exception("No sheet open!");
			}
			//If the cell is empty set its value
			if (currentSheet.Cells[address.Row, address.Column].Value == null) {
				currentSheet.SetValue(address.Address, number);
				Console.WriteLine("Cell[" + address.Address + "] = " + number + " (Was empty)");
				xlsxFile.Save();
				return;
			}
			//if the cell already has a number add the value
			else if (int.TryParse(currentSheet.Cells[address.Row, address.Column].Value.ToString(), out int numberInCell)) {
				currentSheet.SetValue(address.Address, number + numberInCell);
				Console.WriteLine("Cell[" + address.Address + "] = " + (number + numberInCell));
				xlsxFile.Save();
				return;
			}
			Console.WriteLine("Unable to change cell at:" + address.Address + " containing " + currentSheet.GetValue(address.Row, address.Column));
		}

		[Obsolete("Use the generic function 'InsertValue()' to get correctly formatted cells")]
		public void InsertText(ExcelCellAddress address, string text) {
			if (currentSheet == null) {
				throw new Exception("No sheet open!");
			}
			currentSheet.SetValue(address.Address, text);
			Console.WriteLine("Cell[" + address.Address + "] = " + text);
			xlsxFile.Save();
		}


		public void InsertValue<T>(ExcelCellAddress address, T value) {
			if (currentSheet == null) {
				throw new Exception("No sheet open!");
			}
			currentSheet.SetValue(address.Address, value);
			Console.WriteLine("Cell[" + address.Address + "] = " + value.ToString());
			xlsxFile.Save();
		}

		public void MakeANewSpreadsheet(DefinitionParserData data) {
			InsertValue(new ExcelCellAddress(1, 1), "Spreadsheet for enemy: " + currentSheet.Name);
			InsertValue(new ExcelCellAddress(1, 4), "Num killed:");
			InsertValue(new ExcelCellAddress(1, 5), 0);

			nameLookupDictionary = new Dictionary<string, ExcelCellAddress>();
			int[] rowOfEachGroup = new int[data.groups.Length];
			int[] collonOfEachGroup = new int[data.groups.Length];
			int collonOffset = 4;
			int groupcounter = 0;
			foreach (string group in data.groups) {
				rowOfEachGroup[groupcounter] = 2;
				collonOfEachGroup[groupcounter] = groupcounter * collonOffset + 1;
				ExcelCellAddress address = new ExcelCellAddress(rowOfEachGroup[groupcounter], collonOfEachGroup[groupcounter]);
				InsertValue(address, group);
				groupcounter += 1;
			}
			foreach (DefinitionParserData.Entry entry in data.entries) {
				for (int i = 0; i < data.groups.Length; i++) {
					if (data.groups[i] == entry.group) {
						groupcounter = i;
					}
				}
				rowOfEachGroup[groupcounter] += 1;
				InsertValue(new ExcelCellAddress(rowOfEachGroup[groupcounter], collonOfEachGroup[groupcounter]), entry.mainPronounciation);
				InsertValue(new ExcelCellAddress(rowOfEachGroup[groupcounter], collonOfEachGroup[groupcounter] + 1), entry.yangValue);
				InsertValue(new ExcelCellAddress(rowOfEachGroup[groupcounter], collonOfEachGroup[groupcounter] + 2), 0);
				nameLookupDictionary.Add(entry.mainPronounciation, new ExcelCellAddress(rowOfEachGroup[groupcounter], collonOfEachGroup[groupcounter] + 2));
			}
		}

		public void InitialiseWorksheet() {
			if (currentSheet.Cells[1, 1].Value == null) {
				MakeANewSpreadsheet(DefinitionParser.instance.currentGrammarFile);
			}
		}

		public ExcelCellAddress AddressFromName(string name) {
			if (nameLookupDictionary.ContainsKey(name)) {
				return nameLookupDictionary[name];
			}
			else {
				string s = DefinitionParser.instance.currentGrammarFile.GetMainPronounciation(name);
				if (s != null && nameLookupDictionary.ContainsKey(s)) {
					return nameLookupDictionary[s];
				}
			}
			if (!Program.debug) {
				throw new Exception("The word you're looking for isn't in the dictionary due to parsing problems.");
			}
			return new ExcelCellAddress(100, 100);
		}
	}
}
