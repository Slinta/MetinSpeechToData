using Metin2SpeechToData;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using static Metin2SpeechToData.Spreadsheet.SsConstants;


namespace SheetSync {
	class MergeHelper {

		private Dictionary<string, ExcelCellAddress> nameToDropCoutAddress = new Dictionary<string, ExcelCellAddress>();
		private Dictionary<string, Group> nameToGroup = new Dictionary<string, Group>();

		private readonly Dictionary<string, Dictionary<string, ExcelCellAddress>> sheetToNameToDropCountAddress = new Dictionary<string, Dictionary<string, ExcelCellAddress>>();
		private readonly Dictionary<string, Dictionary<string, Group>> sheetToNameToGroupAddress = new Dictionary<string, Dictionary<string, Group>>();

		public MergeHelper(ExcelPackage main, FileInfo[] sessions) {
			List<int> list = new List<int>();

			for (int i = 0; i < sessions.Length; i++) {
				if (sessions[i].Attributes == FileAttributes.Archive) {
					list.Add(i);
				}
			}

			if (list.Count > 0) {
				Console.WriteLine("Found unmerged session files in 'Sessions' folder...");
				foreach (int i in list) {
					if (Confirmation.WrittenConfirmation("Merge '" + sessions[i].Name + "' with main file?")) {
						MergeSession(main, sessions[i]);
						sessions[i].Attributes = FileAttributes.Normal;
					}
				}
			}
		}

		public void MergeSession(ExcelPackage main, FileInfo fileInfo) {
			ExcelPackage package = new ExcelPackage(fileInfo);

			ExcelWorksheet session = package.Workbook.Worksheets["Session"];
			ExcelWorksheet currMain = main.Workbook.Worksheets[H_DEFAULT_SHEET_NAME];

			string currAddress = DATA_FIRST_ENTRY;

			while (session.Cells[currAddress].Value != null) {
				DefinitionParserData.Item item = new DefinitionParserData.Item(
					(string)session.Cells[currAddress].Value,
					new string[0],
					uint.Parse(session.Cells[SpreadsheetHelper.OffsetAddressString(currAddress, 0, 13)].Value.ToString()),
					(string)session.Cells[SpreadsheetHelper.OffsetAddressString(currAddress, 0, 4)].Value
				);
				string enemy = (string)session.Cells[SpreadsheetHelper.OffsetAddress(currAddress, 0, 7).Address].Value;
				string mainSheetName = enemy != UNSPEICIFIED_ENEMY ? enemy : (string)session.Cells[SessionSheet.SESSION_AREA_NAME].Value;

				bool exists = VerifyExistence(main, mainSheetName, enemy != UNSPEICIFIED_ENEMY ? SpreadsheetTemplates.SpreadsheetPresetType.ENEMY : SpreadsheetTemplates.SpreadsheetPresetType.AREA);

				currMain = main.Workbook.Worksheets[mainSheetName];

				if (exists && !sheetToNameToDropCountAddress.ContainsKey(mainSheetName)) {
					Dictionaries d = LoadSpreadsheet(main.Workbook.Worksheets[mainSheetName], enemy != UNSPEICIFIED_ENEMY ? SpreadsheetTemplates.SpreadsheetPresetType.ENEMY : SpreadsheetTemplates.SpreadsheetPresetType.AREA);
					sheetToNameToDropCountAddress.Add(mainSheetName, d.addresses);
					sheetToNameToGroupAddress.Add(mainSheetName, d.groups);
					nameToDropCoutAddress = sheetToNameToDropCountAddress[mainSheetName];
					nameToGroup = sheetToNameToGroupAddress[mainSheetName];
				}
				else {
					nameToDropCoutAddress = sheetToNameToDropCountAddress[mainSheetName];
					nameToGroup = sheetToNameToGroupAddress[mainSheetName];
				}



				if (nameToDropCoutAddress.ContainsKey(item.mainPronounciation)) {
					currMain.SetValue(nameToDropCoutAddress[item.mainPronounciation].Address, item.mainPronounciation);
					currMain.SetValue(SpreadsheetHelper.OffsetAddressString(nameToDropCoutAddress[item.mainPronounciation].Address,0,1), item.yangValue);
					currMain.SetValue(SpreadsheetHelper.OffsetAddressString(nameToDropCoutAddress[item.mainPronounciation].Address,0,2), 
						(int)currMain.Cells[SpreadsheetHelper.OffsetAddressString(nameToDropCoutAddress[item.mainPronounciation].Address, 0, 2)].Value + 1);
				}
				else {
					AddItemEntry(currMain, item);
				}



				currAddress = SpreadsheetHelper.OffsetAddressString(currAddress, 1, 0);
			}
			main.Save();
		}

		private bool VerifyExistence(ExcelPackage main, string sheetName, SpreadsheetTemplates.SpreadsheetPresetType preset) {
			for (int i = 0; i < main.Workbook.Worksheets.Count; i++) {
				if (main.Workbook.Worksheets[i].Name == sheetName) {
					return true;
				}
			}
			SpreadsheetTemplates t = new SpreadsheetTemplates();
			ExcelWorksheets sheets = t.LoadTemplates();
			switch (preset) {
				case SpreadsheetTemplates.SpreadsheetPresetType.AREA: {
					main.Workbook.Worksheets.Add(sheetName, sheets["Area"]);
					break;
				}
				case SpreadsheetTemplates.SpreadsheetPresetType.ENEMY: {
					main.Workbook.Worksheets.Add(sheetName, sheets["Enemy"]);
					break;
				}
			}
			t.Dispose();
			main.Save();
			return false;
		}

		private void AddItemEntry(ExcelWorksheet currentSheet, DefinitionParserData.Item entry) {
			ExcelCellAddress current = new ExcelCellAddress(ITEM_ROW, GROUP_COL + GetGroupIndex(currentSheet, entry.group) * H_COLUMN_INCREMENT);

			//Add the group if it does't exist
			if (!currentSheet.Cells[current.Row - 1, current.Column, current.Row - 1, current.Column + 2].Merge) {
				currentSheet.Cells[current.Row - 1, current.Column, current.Row - 1, current.Column + 2].Merge = true;
			}
			currentSheet.Cells[current.Row - 1, current.Column, current.Row - 1, current.Column + 2].Value = entry.group;
			currentSheet.Cells[current.Row - 1, current.Column, current.Row - 1, current.Column + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

			//Find free spot in the group
			while (currentSheet.GetValue(current.Row, current.Column) != null) {
				current = SpreadsheetHelper.OffsetAddress(current, 1, 0);
			}

			//Insert
			currentSheet.SetValue(current.Address, entry.mainPronounciation);
			currentSheet.SetValue(SpreadsheetHelper.OffsetAddress(current, 0, 1).Address, entry.yangValue);
			currentSheet.SetValue(SpreadsheetHelper.OffsetAddress(current, 0, 2).Address, 1);
			nameToDropCoutAddress.Add(entry.mainPronounciation, new ExcelCellAddress(current.Row, current.Column + 2));
		}

		private short GetGroupIndex(ExcelWorksheet sheet, string group) {
			ExcelCellAddress addr = new ExcelCellAddress(GROUP_ROW, GROUP_COL);
			string val = sheet.GetValue<string>(addr.Row, addr.Column);
			short counter = 0;
			while (val != group && val != "" && val != "GROUP_NAME") {
				addr = SpreadsheetHelper.OffsetAddress(addr, 0, H_COLUMN_INCREMENT);
				val = sheet.GetValue<string>(addr.Row, addr.Column);
				counter++;
			}
			return counter;
		}


		/// <summary>
		/// Parses spreadsheet
		/// </summary>
		public Dictionaries LoadSpreadsheet(ExcelWorksheet sheet, SpreadsheetTemplates.SpreadsheetPresetType type) {
			Dictionaries d = new Dictionaries {
				addresses = new Dictionary<string, ExcelCellAddress>(),
				groups = new Dictionary<string, Group>()
			};

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
						Group g = new Group(address, new ExcelCellAddress(address.Row + 1, address.Column));
						d.groups.Add(group, g);
					}
					foreach (DefinitionParserData.Item entry in data.entries) {
						Group g = d.groups[entry.group];
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
						current = SpreadsheetHelper.Advance(sheet, current);
					}
					return d;
				}
				default: {
					throw new CustomException("Uncathegorized sheet entered!");
				}
			}
		}


		public struct Dictionaries {
			public Dictionary<string, ExcelCellAddress> addresses;
			public Dictionary<string, Group> groups;
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
	}
}
