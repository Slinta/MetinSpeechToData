using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using static Metin2SpeechToData.Configuration;
using Metin2SpeechToData.Structures;

namespace Metin2SpeechToData {
	public class SpreadsheetInteraction : IDisposable {

		private ExcelPackage xlsxFile;
		private ExcelWorkbook content;

		private readonly SpreadsheetTemplates templates;
		private readonly SpreadsheetHelper helper;
		private ExcelWorksheet _currentSheet;
		public ExcelWorksheet mainSheet { get; }

		private Dictionary<string, ExcelCellAddress> currentNameToConutAddress;
		private Dictionary<string, Group> currentGroupsByName;

		private readonly Dictionary<string, Dictionary<string, ExcelCellAddress>> sheetToAdresses;
		private readonly Dictionary<string, Dictionary<string, Group>> sheetToGroups;

		private uint currModificationsToXlsx = 0;

		#region Constructor
		/// <summary>
		/// Initializes spreadsheet control with given file
		/// </summary>
		public SpreadsheetInteraction(FileInfo path) {
			xlsxFile = new ExcelPackage(path);
			content = xlsxFile.Workbook;
			_currentSheet = content.Worksheets["Metin2 Drop Analyzer"];
			mainSheet = currentSheet;
			helper = new SpreadsheetHelper(this);

			templates = new SpreadsheetTemplates(this);
			sheetToAdresses = new Dictionary<string, Dictionary<string, ExcelCellAddress>>();
			sheetToGroups = new Dictionary<string, Dictionary<string, Group>>();

			//TODO this testing stuff
			//SpreadsheetHelper.HyperlinkCell(currentSheet, "Q1", content.Worksheets[1], "B2", "Link to B2");
			//templates.InitAreaSheet(content, new DefinitionParserData("AA", new string[3] { "AAA", "BBBB", "CCC" }, new DefinitionParserData.Item[2] {
			//	new DefinitionParserData.Item("AAAAAAAA", new string[0],1000,"AAA"),
			//	new DefinitionParserData.Item("CSCSCSCSC", new string[0], 1000000, "CCC"),
			//}));
			//Save();
		}
		#endregion

		public void SwitchFile(FileInfo newFile) {
			Save();

			if (newFile == Program.config.xlsxFile) {
				Console.WriteLine("Switching back to default file");
				xlsxFile = new ExcelPackage(Program.config.xlsxFile);
				content = xlsxFile.Workbook;
				currentSheet = content.Worksheets[1];
				return;
			}

			if (!File.Exists(newFile.FullName)) {
				Console.WriteLine("Creating new file for session.");
				if (!string.IsNullOrWhiteSpace(newFile.FullName)) {
					File.Create(newFile.FullName);
				}
				else {
					newFile = new FileInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + Path.GetRandomFileName() + ".xlsx");
					newFile.Attributes = FileAttributes.Hidden | FileAttributes.Archive;
					File.Create(newFile.FullName);
				}
			}
			currentSheet.Dispose();
			content.Dispose();
			xlsxFile.Dispose();

			xlsxFile = new ExcelPackage(newFile);
			content = xlsxFile.Workbook;
			if (content.Worksheets["Session"] == null) {
				content.Worksheets.Add("Session");
				currentSheet = content.Worksheets["Session"];
			}
		}


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
				Dicts dicts = SpreadsheetHelper.LoadSpreadsheet(content, sheetName, helper.GetSheetType(sheetName));
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
			int numberInCell = 0;
			//If the cell is empty set its value
			if (_currentSheet.Cells[address.Row, address.Column].Value == null) {
				_currentSheet.SetValue(address.Address, number);
			}
			//if the cell already has a number add the value
			else if (int.TryParse(_currentSheet.Cells[address.Row, address.Column].Value.ToString(), out numberInCell)) {
				_currentSheet.SetValue(address.Address, number + numberInCell);
			}
			else {
				Console.WriteLine("Unable to change cell at:" + address.Address + " containing " + _currentSheet.GetValue(address.Row, address.Column));
				return;
			}
			currModificationsToXlsx++;
			Console.WriteLine("Cell[" + address.Address + "] = " + (number + numberInCell));
			Save();
		}

		/// <summary>
		/// Insert 'value' into cell at 'address' in current worksheet
		/// </summary>
		public void InsertValue<T>(string address, T value) {
			InsertValue(new ExcelCellAddress(address), value);
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
			ExcelWorksheet newSheet = templates.InitAreaSheet(content,DefinitionParser.instance.currentGrammarFile);
			Dicts dicts = SpreadsheetHelper.LoadSpreadsheet(newSheet, SpreadsheetTemplates.SpreadsheetPresetType.AREA);
			if (!sheetToAdresses.ContainsKey(areaName)) {
				sheetToAdresses.Add(areaName, dicts.addresses);
				sheetToGroups.Add(areaName, dicts.groups);
			}
			_currentSheet = content.Worksheets[areaName];
			Save();
		}

		public void InitMobSheet(string mobName) {
			ExcelWorksheet newSheet = templates.InitEnemySheet(content, mobName, Program.gameRecognizer.enemyHandling.mobDrops);
			Dicts dicts = SpreadsheetHelper.LoadSpreadsheet(newSheet, SpreadsheetTemplates.SpreadsheetPresetType.ENEMY);
			if (!sheetToAdresses.ContainsKey(mobName)) {
				sheetToAdresses.Add(mobName, dicts.addresses);
				sheetToGroups.Add(mobName, dicts.groups);
			}
			_currentSheet = content.Worksheets[mobName];
			Save();
		}
		#endregion


		/// <summary>
		/// Dynamically append new entries to the current sheet
		/// </summary>
		[Obsolete("FIX,TIHS",true)]
		public void AddItemEntryToCurrentSheet(DefinitionParserData.Item entry) {

			ExcelCellAddress current = new ExcelCellAddress(2, 1 + Program.gameRecognizer.enemyHandling.mobDrops.GetGroupNumberForEnemy(_currentSheet.Name, entry.group) * 4);
			int maxDetph = 10;

			//Add the group if it does't exist
			if (!_currentSheet.Cells[current.Row, current.Column, current.Row, current.Column + 2].Merge) {
				_currentSheet.Cells[current.Row, current.Column, current.Row, current.Column + 2].Merge = true;
			}
			_currentSheet.Cells[current.Row, current.Column, current.Row, current.Column + 2].Value = entry.group;
			_currentSheet.Cells[current.Row, current.Column, current.Row, current.Column + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

			//Find free spot in the group
			while (_currentSheet.Cells[current.Address].Value != null) {
				current = new ExcelCellAddress(current.Row + 1, current.Column);
				if (current.Row >= maxDetph) {
					current = new ExcelCellAddress(2, current.Column + 4);
				}
			}

			//Insert
			InsertValue(current, entry.mainPronounciation);
			InsertValue(new ExcelCellAddress(current.Row, current.Column + 1), entry.yangValue);
			InsertValue(new ExcelCellAddress(current.Row, current.Column + 2), 0);
			sheetToAdresses[_currentSheet.Name].Add(entry.mainPronounciation, new ExcelCellAddress(current.Row, current.Column + 2));
			Save();
		}

		/// <summary>
		/// Dynamically remove new entries from the current sheet
		/// </summary>
		[Obsolete("FIX,TIHS", true)]
		public void RemoveItemEntryFromCurrentSheet(DefinitionParserData.Item entry) {
			ExcelCellAddress current = new ExcelCellAddress("A2");
			int maxDetph = 10;

			while (_currentSheet.Cells[current.Address].GetValue<string>() != entry.mainPronounciation) {
				current = new ExcelCellAddress(current.Row + 1, current.Column);
				if (current.Row >= maxDetph) {
					current = new ExcelCellAddress(2, current.Column + 4);
				}
			}
			//Found the cell
			InsertValue<object>(current, null);
			InsertValue<object>(new ExcelCellAddress(current.Row, current.Column + 1), null);
			InsertValue<object>(new ExcelCellAddress(current.Row, current.Column + 2), null);
			sheetToAdresses[_currentSheet.Name].Remove(entry.mainPronounciation);
		}

		/// <summary>
		/// Adds a new sheet to current workbook as 'type' spreadsheet
		/// </summary>
		public ExcelWorksheet AddSheet(string name, SpreadsheetTemplates.SpreadsheetPresetType type) {
			if(content.Worksheets[name] != null) {
				throw new CustomException("File already contains sheet named " + name);
			}
			ExcelWorksheets reference = templates.LoadTemplates();
			ExcelWorksheet selected = null;
			switch (type) {
				case SpreadsheetTemplates.SpreadsheetPresetType.MAIN: {
					throw new CustomException("Tried to add a main sheet to a file, this is unsupported");
				}
				case SpreadsheetTemplates.SpreadsheetPresetType.AREA: {
					selected = reference["Area"];
					break;
				}
				case SpreadsheetTemplates.SpreadsheetPresetType.ENEMY: {
					selected = reference["Enemy"];
					break;
				}
				case SpreadsheetTemplates.SpreadsheetPresetType.SESSION: {
					selected = reference["Session"];
					break;
				}
				default: {
					throw new CustomException("Invalid Preset type!");
				}
			}
			content.Worksheets.Add(name, selected);
			templates.Dispose();
			selected.Dispose();
			reference.Dispose();
			return content.Worksheets[name];
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
		/// Gets group stuct with 'identifier' name
		/// </summary>
		public Group GetGroup(string identifier) {
			try {
				return currentGroupsByName[identifier];
			}
			catch {
				throw new CustomException("Item not present in currentNameToConutAddress!");
			}
		}

		/// <summary>
		/// Saves current changes to the .xlsx file
		/// </summary>
		public void Save() {
			try {
				currModificationsToXlsx++;
				if (currModificationsToXlsx == sheetChangesBeforeSaving) {
					xlsxFile.Save();
					currModificationsToXlsx = 0;
				}
			}
			catch (Exception e) {
				Console.WriteLine("Could not update. Is the file already open ?\n" + e.Message);
			}
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

		public struct Group {

			public Group(ExcelCellAddress groupName, ExcelCellAddress elementNameFirstIndex) {
				this.groupName = groupName;
				this.elementNameFirstIndex = elementNameFirstIndex;
				this.totalEntries = 0;
			}

			public ExcelCellAddress groupName { get; }
			public ExcelCellAddress elementNameFirstIndex { get; }
			public int totalEntries { get; set; }
		}

		#region IDisposable Support
		private bool disposedValue = false;

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				Save();
				if (disposing) {
					return;
				}
				xlsxFile.Dispose();
				content.Dispose();
				currentSheet.Dispose();
				disposedValue = true;
			}
		}

		public void Dispose() {
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
