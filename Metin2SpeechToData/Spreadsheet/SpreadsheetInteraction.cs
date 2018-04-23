using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public class SpreadsheetInteraction {

		private ExcelPackage xlsxFile;
		private ExcelWorkbook content;

		private SpreadsheetTemplates templates;
		private SpreadsheetHelper helper;
		private ExcelWorksheet _currentSheet;
		private Dictionary<string, ExcelCellAddress> currentNameToConutAddress;
		private Dictionary<string, Group> currentGroupsByName;

		private Dictionary<string, Dictionary<string, ExcelCellAddress>> sheetToAdresses = new Dictionary<string, Dictionary<string, ExcelCellAddress>>();
		private Dictionary<string, Dictionary<string, Group>> sheetToGroups = new Dictionary<string, Dictionary<string, Group>>();


		/// <summary>
		/// Initializes spreadsheet control with given file
		/// </summary>
		public SpreadsheetInteraction(FileInfo path) {
			xlsxFile = new ExcelPackage(path);
			content = xlsxFile.Workbook;
			_currentSheet = content.Worksheets["Metin2 Drop Analyzer"];

			helper = new SpreadsheetHelper(this);

			templates = new SpreadsheetTemplates(this);
			templates.InitializeMainSheet();
		}

		/// <summary>
		/// Open existing worksheet or add one if it does't exist
		/// </summary>
		public void OpenWorksheet(string sheetName) {
			if (content.Worksheets[sheetName] == null) {
				content.Worksheets.Add(sheetName);
				SpreadsheetTemplates.SpreadsheetPresetType type = helper.GetSheetType(sheetName);
				switch (type) {
					case SpreadsheetTemplates.SpreadsheetPresetType.AREA: {
						InitAreaSheet(sheetName);
						currentSheet = content.Worksheets[sheetName];
						AutoAdjustColumns();
						return;
					}
					case SpreadsheetTemplates.SpreadsheetPresetType.ENEMY: {
						InitMobSheet(sheetName);
						currentSheet = content.Worksheets[sheetName];
						return;
					}
				}
			}
			if (!sheetToAdresses.ContainsKey(sheetName)) {
				SpreadsheetHelper.Dicts dicts = SpreadsheetHelper.LoadSpreadsheet(content, sheetName, helper.GetSheetType(sheetName));
				sheetToAdresses.Add(sheetName, dicts.addresses);
				sheetToGroups.Add(sheetName, dicts.groups);
			}
			currentSheet = content.Worksheets[sheetName];
		}


		/// <summary>
		/// Adds a 'number' to the number cell at 'address' and saves the document
		/// </summary>
		public void AddNumberTo(ExcelCellAddress address, int number) {
			if (_currentSheet == null) {
				throw new CustomException("No sheet open!");
			}
			//If the cell is empty set its value
			if (_currentSheet.Cells[address.Row, address.Column].Value == null) {
				_currentSheet.SetValue(address.Address, number);
				Console.WriteLine("Cell[" + address.Address + "] = " + number + " (Was empty)");
				xlsxFile.Save();
				return;
			}
			//if the cell already has a number add the value
			else if (int.TryParse(_currentSheet.Cells[address.Row, address.Column].Value.ToString(), out int numberInCell)) {
				_currentSheet.SetValue(address.Address, number + numberInCell);
				Console.WriteLine("Cell[" + address.Address + "] = " + (number + numberInCell));
				xlsxFile.Save();
				return;
			}
			Console.WriteLine("Unable to change cell at:" + address.Address + " containing " + _currentSheet.GetValue(address.Row, address.Column));
		}

		/// <summary>
		/// Insert 'value' into cell at 'address' in current worksheet
		/// </summary>
		public void InsertValue<T>(ExcelCellAddress address, T value) {
			if (_currentSheet == null) {
				throw new CustomException("No sheet open!");
			}
			_currentSheet.SetValue(address.Address, value);
			Console.WriteLine("Cell[" + address.Address + "] = " + value.ToString());
			xlsxFile.Save();
		}

		private void AutoAdjustColumns() {
			double currMaxWidth = 0;
			foreach (Group g in currentGroupsByName.Values) {
				for (int i = 0; i < g.totalEntries; i++) {
					_currentSheet.Cells[g.elementNameFirstIndex.Row + i, g.elementNameFirstIndex.Column].AutoFitColumns();
					if (_currentSheet.Column(g.elementNameFirstIndex.Column).Width >= currMaxWidth) {
						currMaxWidth = _currentSheet.Column(g.elementNameFirstIndex.Column).Width;
					}
				}

				_currentSheet.Column(g.elementNameFirstIndex.Column).Width = currMaxWidth;
				currMaxWidth = 0;

				for (int i = 0; i < g.totalEntries; i++) {
					int s = _currentSheet.GetValue<int>(g.yangValueFirstIndex.Row + i, g.yangValueFirstIndex.Column);
					_currentSheet.Column(g.yangValueFirstIndex.Column).Width = SpreadsheetConstants.GetCellWidth(s, false);
					if (_currentSheet.Column(g.yangValueFirstIndex.Column).Width > currMaxWidth) {
						currMaxWidth = _currentSheet.Column(g.yangValueFirstIndex.Column).Width;
					}
				}

				_currentSheet.Column(g.yangValueFirstIndex.Column).Width = currMaxWidth;
				currMaxWidth = 0;
			}
		}

		public void InitAreaSheet(string areaName) {
			_currentSheet = content.Worksheets[areaName];
			SpreadsheetHelper.Dicts dicts = templates.InitializeAreaSheet(DefinitionParser.instance.currentGrammarFile);
			if (!sheetToAdresses.ContainsKey(areaName)) {
				sheetToAdresses.Add(areaName, dicts.addresses);
				sheetToGroups.Add(areaName, dicts.groups);
			}
			Save();
		}

		public void InitMobSheet(string mobName) {
			_currentSheet = content.Worksheets[mobName];
			SpreadsheetHelper.Dicts dicts = templates.InitializeMobSheet(mobName, Program.enemyHandling.mobDrops);
			if (!sheetToAdresses.ContainsKey(mobName)) {
				sheetToAdresses.Add(mobName, dicts.addresses);
				sheetToGroups.Add(mobName, dicts.groups);
			}
			Save();
		}

		public void AddItemEntryToCurrentSheet(string itemName) {
			templates.AddItemEntry(_currentSheet, new DefinitionParserData.Entry {
				mainPronounciation = itemName,
				yangValue = DefinitionParser.instance.currentGrammarFile.GetYangValue(itemName),
				group = null,
				ambiguous = null
			});
		}

		public ExcelCellAddress AddressFromName(string name) {
			if (currentNameToConutAddress.ContainsKey(name)) {
				return currentNameToConutAddress[name];
			}
			if (!Program.debug) {
				throw new CustomException("The word you're looking for isn't in the dictionary due to parsing problems.");
			}
			else {
				return new ExcelCellAddress(100, 100);
			}
		}

		public void Save() {
			xlsxFile.Save();
		}


		public ExcelWorksheet currentSheet {
			get { return _currentSheet; }
			private set {
				currentNameToConutAddress = sheetToAdresses[value.Name];
				currentGroupsByName = sheetToGroups[value.Name];
				_currentSheet = value;
			}
		}

		public void AddSheetToAddressEntry(string sheetName, string itemName, ExcelCellAddress address) {
			sheetToAdresses[sheetName].Add(itemName, address);
		}

		public struct Group {
			public ExcelCellAddress groupName;
			public ExcelCellAddress elementNameFirstIndex;
			public ExcelCellAddress yangValueFirstIndex;
			public ExcelCellAddress totalCollectedFirstIndex;
			public int totalEntries;
		}
	}
}
