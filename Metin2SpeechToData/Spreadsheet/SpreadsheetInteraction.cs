using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Metin2SpeechToData.Structures;

using static Metin2SpeechToData.Configuration;
using static Metin2SpeechToData.Spreadsheet.SsConstants;

namespace Metin2SpeechToData {
	public class SpreadsheetInteraction : IDisposable {

		private readonly ExcelPackage xlsxFile;
		private readonly ExcelWorkbook content;

		private readonly SpreadsheetTemplates templates;

		public ExcelWorksheet mainSheet { get; }

		private Dictionary<string, ExcelCellAddress> currentNameToConutAddress;
		private Dictionary<string, Group> currentGroupsByName;

		private uint currModificationsToXlsx = 0;


		/// <summary>
		/// Accessor to currently open sheet, SET: automatically pick sheet addresses and groups
		/// </summary>
		//TODO: get rid of this reference eveywhere, this will be used only during mergeing
		private ExcelWorksheet currentSheet { get; set; }

		/// <summary>
		/// Current session data and file
		/// </summary>
		public SessionSheet currentSession { get; set; }

		#region Constructor
		/// <summary>
		/// Initializes spreadsheet control with given file
		/// </summary>
		public SpreadsheetInteraction(FileInfo path) {
			xlsxFile = new ExcelPackage(path);
			content = xlsxFile.Workbook;
			currentSheet = content.Worksheets["Metin2 Drop Analyzer"];
			mainSheet = currentSheet;
			templates = new SpreadsheetTemplates(this);

			//TODO this testing stuff
			//SpreadsheetHelper.HyperlinkCell(currentSheet, "Q1", content.Worksheets[1], "B2", "Link to B2");
			//templates.InitAreaSheet(content, new DefinitionParserData("AA", new string[3] { "AAA", "BBBB", "CCC" }, new DefinitionParserData.Item[2] {
			//	new DefinitionParserData.Item("AAAAAAAA", new string[0],1000,"AAA"),
			//	new DefinitionParserData.Item("CSCSCSCSC", new string[0], 1000000, "CCC"),
			//}));
			//Save();
		}
		#endregion

		public void StartSession(string grammar) {
			if(currentSession != null) {
				currentSession.Finish();
				currentSession = new SessionSheet(grammar);
			}
			else {
				currentSession = new SessionSheet(grammar);
			}
		}

		/// <summary>
		/// Adds a 'number' to the number cell at 'address' and saves the document
		/// </summary>
		public void AddNumberTo(ExcelCellAddress address, int number) {
			if (currentSheet == null) {
				throw new CustomException("No sheet open!");
			}
			int numberInCell = 0;
			//If the cell is empty set its value
			if (currentSheet.Cells[address.Row, address.Column].Value == null) {
				currentSheet.SetValue(address.Address, number);
			}
			//if the cell already has a number add the value
			else if (int.TryParse(currentSheet.Cells[address.Row, address.Column].Value.ToString(), out numberInCell)) {
				currentSheet.SetValue(address.Address, number + numberInCell);
			}
			else {
				Console.WriteLine("Unable to change cell at:" + address.Address + " containing " + currentSheet.GetValue(address.Row, address.Column));
				return;
			}
			currModificationsToXlsx++;
			Console.WriteLine("Cell[" + address.Address + "] = " + (number + numberInCell));
			Save();
		}

		/// <summary>
		/// Insert 'value' into cell at 'address' in current worksheet
		/// </summary>
		public void InsertValue<T>(ExcelCellAddress address, T value) {
			if (currentSheet == null) {
				throw new CustomException("No sheet open!");
			}
			currentSheet.SetValue(address.Address, value);
			Console.WriteLine("Cell[" + address.Address + "] = " + value);
			Save();
		}


		#region Sheet initializers
		public void InitAreaSheet(string areaName) {
			ExcelWorksheet newSheet = templates.InitAreaSheet(content, DefinitionParser.instance.currentGrammarFile);
			Dicts dicts = SpreadsheetHelper.LoadSpreadsheet(newSheet, SpreadsheetTemplates.SpreadsheetPresetType.AREA);
			currentNameToConutAddress = dicts.addresses;
			currentGroupsByName = dicts.groups;
			currentSheet = content.Worksheets[areaName];
			Save();
		}

		public void InitMobSheet(string mobName) {
			ExcelWorksheet newSheet = templates.InitEnemySheet(content, mobName, Program.gameRecognizer.enemyHandling.mobDrops);
			Dicts dicts = SpreadsheetHelper.LoadSpreadsheet(newSheet, SpreadsheetTemplates.SpreadsheetPresetType.ENEMY);
			currentNameToConutAddress = dicts.addresses;
			currentGroupsByName = dicts.groups;
			currentSheet = content.Worksheets[mobName];
			Save();
		}
		#endregion


		/// <summary>
		/// Dynamically append new entries to the current sheet
		/// </summary>
		public void AddItemEntryToCurrentSheet(DefinitionParserData.Item entry) {
			ExcelCellAddress current = new ExcelCellAddress(ITEM_ROW, GROUP_COL + Program.gameRecognizer.enemyHandling.mobDrops.GetGroupNumberForEnemy(currentSheet.Name, entry.group) * H_COLUMN_INCREMENT);

			//Add the group if it does't exist
			if (!currentSheet.Cells[current.Row, current.Column, current.Row, current.Column + 2].Merge) {
				currentSheet.Cells[current.Row, current.Column, current.Row, current.Column + 2].Merge = true;
			}
			currentSheet.Cells[current.Row, current.Column, current.Row, current.Column + 2].Value = entry.group;
			currentSheet.Cells[current.Row, current.Column, current.Row, current.Column + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

			//Find free spot in the group
			while (currentSheet.GetValue(current.Row, current.Column) != null) {
				current = SpreadsheetHelper.Advance(currentSheet, current);
			}

			//Insert
			InsertValue(current, entry.mainPronounciation);
			InsertValue(SpreadsheetHelper.OffsetAddress(current, 0, 1), entry.yangValue);
			InsertValue(SpreadsheetHelper.OffsetAddress(current, 0, 1), 0);
			currentNameToConutAddress.Add(entry.mainPronounciation, new ExcelCellAddress(current.Row, current.Column + 2));
			Save();
		}

		/// <summary>
		/// Dynamically remove new entries from the current sheet
		/// </summary>
		public void RemoveItemEntryFromCurrentSheet(DefinitionParserData.Item entry) {
			ExcelCellAddress current = new ExcelCellAddress(ITEM_ROW, GROUP_COL + Program.gameRecognizer.enemyHandling.mobDrops.GetGroupNumberForEnemy(currentSheet.Name, entry.group) * H_COLUMN_INCREMENT);

			while (currentSheet.Cells[current.Address].GetValue<string>() != entry.mainPronounciation) {
				current = SpreadsheetHelper.Advance(currentSheet, current);
			}

			InsertValue<object>(current, null);
			InsertValue<object>(SpreadsheetHelper.OffsetAddress(current, 0, 1), null);
			InsertValue<object>(SpreadsheetHelper.OffsetAddress(current, 0, 2), null);
			currentNameToConutAddress.Remove(entry.mainPronounciation);
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
