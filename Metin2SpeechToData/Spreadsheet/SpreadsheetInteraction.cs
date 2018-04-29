using System;
using System.Linq;
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

		#region Constructor / Destructor
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

		~SpreadsheetInteraction() {
			//TODO make this work somehow, and reduce the amount of saving during the program run cycle
			Save();
		}
		#endregion

		/// <summary>
		/// Open existing worksheet or add one and populate it if the sheet does not exist
		/// </summary>
		public void OpenWorksheet(string sheetName) {
			if (content.Worksheets[sheetName] == null) {
				content.Worksheets.Add(sheetName);
				SpreadsheetTemplates.SpreadsheetPresetType type = helper.GetSheetType(sheetName);
				switch (type) {
					case SpreadsheetTemplates.SpreadsheetPresetType.AREA: {
						InitAreaSheet(sheetName);
						currentSheet = content.Worksheets[sheetName];
						helper.AutoAdjustColumns(currentGroupsByName.Values);
						return;
					}
					case SpreadsheetTemplates.SpreadsheetPresetType.ENEMY: {
						InitMobSheet(sheetName);
						currentSheet = content.Worksheets[sheetName];
						helper.AutoAdjustColumns(currentGroupsByName.Values);
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
				Save();
				return;
			}
			//if the cell already has a number add the value
			else if (int.TryParse(_currentSheet.Cells[address.Row, address.Column].Value.ToString(), out int numberInCell)) {
				_currentSheet.SetValue(address.Address, number + numberInCell);
				Console.WriteLine("Cell[" + address.Address + "] = " + (number + numberInCell));
				Save();
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
			Console.WriteLine("Cell[" + address.Address + "] = " + value);
			Save();
		}

		#region Sheet initializers
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
		#endregion

		/// <summary>
		/// Dynamically append new entries to the current sheet
		/// </summary>
		public void AddItemEntryToCurrentSheet(string itemName) {
			templates.AddItemEntry(_currentSheet, new DefinitionParserData.Entry {
				mainPronounciation = itemName,
				yangValue = DefinitionParser.instance.currentGrammarFile.GetYangValue(itemName),
				group = DefinitionParser.instance.currentGrammarFile.GetGroup(itemName),
				ambiguous = null
			});
		}

		/// <summary>
		/// Dynamically remove new entries from the current sheet
		/// </summary>
		public void RemoveItemEntryFromCurrentSheet(string itemName) {
			templates.RemoveItemEntry(_currentSheet, new DefinitionParserData.Entry {
				mainPronounciation = itemName,
				yangValue = DefinitionParser.instance.currentGrammarFile.GetYangValue(itemName),
				group = DefinitionParser.instance.currentGrammarFile.GetGroup(itemName),
				ambiguous = null
			});
		}

		/// <summary>
		/// Gets the item address belonging to 'itemName' in current sheet
		/// </summary>
		public ExcelCellAddress GetAddress(string itemName) {
			try {
				return currentNameToConutAddress[itemName];
			}
			catch {
				throw new CustomException("Item not present in currentNameToConutAddress!");
			}
		}

		/// <summary>
		/// Saves current changes to the .xlsx file
		/// </summary>
		public void Save() {
			//TODO fix this crashing after quitting
			xlsxFile.Save();
		}

		/// <summary>
		/// Accessor to currently open sheet, SET: automatically pick sheet addresses and groups
		/// </summary>
		public ExcelWorksheet currentSheet {
			get { return _currentSheet; }
			private set {
				currentNameToConutAddress = sheetToAdresses[value.Name];
				currentGroupsByName = sheetToGroups[value.Name];
				_currentSheet = value;
			}
		}

		/// <summary>
		/// Wrapper to add entry 'address' indexed as sheet name 'itemName' to sheet 'sheetName' 
		/// </summary>
		public void AddSheetToAddressEntry(string sheetName, string itemName, ExcelCellAddress address) {
			sheetToAdresses[sheetName].Add(itemName, address);
		}

		/// <summary>
		/// Wrapper to remove entry 'address' indexed as sheet name 'itemName' to sheet 'sheetName' 
		/// </summary>
		public void RemoveSheetToAddressEntry(string sheetName, string itemName) {
			sheetToAdresses[sheetName].Remove(itemName);
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
