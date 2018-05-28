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

			string currAddress = DATA_FIRST_ENTRY;

			while (session.Cells[currAddress].Value != null) {
				DefinitionParserData.Item item = new DefinitionParserData.Item(
					(string)session.Cells[currAddress].Value,
					new string[0],
					uint.Parse(session.Cells[SpreadsheetHelper.OffsetAddressString(currAddress, 0, 13)].Value.ToString()),
					(string)session.Cells[SpreadsheetHelper.OffsetAddressString(currAddress, 0, 4)].Value
				);
				string enemy = (string)session.Cells[SpreadsheetHelper.OffsetAddress(currAddress, 0, 7).Address].Value;
				string currentSheetName = enemy != UNSPEICIFIED_ENEMY ? enemy : (string)session.Cells[SessionSheet.SESSION_AREA_NAME].Value;

				bool exists = VerifyExistence(main, currentSheetName);
				ExcelWorksheet currMain;
				if (!exists) {
					SpreadsheetTemplates t = new SpreadsheetTemplates();
					SpreadsheetTemplates.SpreadsheetPresetType type = (enemy == currentSheetName ? SpreadsheetTemplates.SpreadsheetPresetType.ENEMY : SpreadsheetTemplates.SpreadsheetPresetType.AREA);
					currMain = t.CreateFromTemplate(main.Workbook, type, currentSheetName);
					if (type == SpreadsheetTemplates.SpreadsheetPresetType.AREA) {
						InitAreaSheet(currentSheetName, currMain);
					}
					else {
						InitMobSheet(currentSheetName, (string)session.Cells[SessionSheet.SESSION_AREA_NAME].Value, currMain);
					}
					Console.WriteLine("Creating new sheet for " + currentSheetName);
					main.Save();
				}
				else {
					currMain = main.Workbook.Worksheets[currentSheetName];
				}


				if (!sheetToNameToDropCountAddress.ContainsKey(currentSheetName)) {
					Dictionaries d = LoadSpreadsheet(main.Workbook.Worksheets[currentSheetName], enemy != UNSPEICIFIED_ENEMY ? SpreadsheetTemplates.SpreadsheetPresetType.ENEMY : SpreadsheetTemplates.SpreadsheetPresetType.AREA);
					sheetToNameToDropCountAddress.Add(currentSheetName, d.addresses);
					sheetToNameToGroupAddress.Add(currentSheetName, d.groups);
					nameToDropCoutAddress = sheetToNameToDropCountAddress[currentSheetName];
					nameToGroup = sheetToNameToGroupAddress[currentSheetName];
				}
				else {
					nameToDropCoutAddress = sheetToNameToDropCountAddress[currentSheetName];
					nameToGroup = sheetToNameToGroupAddress[currentSheetName];
				}




				if (nameToDropCoutAddress.ContainsKey(item.mainPronounciation)) {
					int currValue = (int)currMain.Cells[nameToDropCoutAddress[item.mainPronounciation].Row, nameToDropCoutAddress[item.mainPronounciation].Column].Value;
					currMain.SetValue(nameToDropCoutAddress[item.mainPronounciation].Address, currValue + 1);
				}
				else {
					AddItemEntry(currMain, item);
				}



				currAddress = SpreadsheetHelper.OffsetAddressString(currAddress, 1, 0);
			}
			main.Save();
		}

		private bool VerifyExistence(ExcelPackage main, string sheetName) {
			for (int i = 1; i <= main.Workbook.Worksheets.Count; i++) {
				if (main.Workbook.Worksheets[i].Name == sheetName) {
					return true;
				}
			}
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

		private byte GetGroupIndex(ExcelWorksheet sheet, string group) {
			ExcelCellAddress addr = new ExcelCellAddress(GROUP_ROW, GROUP_COL);
			string val = sheet.GetValue<string>(addr.Row, addr.Column);
			byte counter = 0;
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
					int totalGroups = CountGroups(sheet);
					byte[] rowOfEachGroup = new byte[totalGroups];
					byte[] columnOfEachGroup = new byte[totalGroups];
					byte groupcounter = 0;

					for (int i = 0; i < totalGroups; i++) {
						rowOfEachGroup[groupcounter] = 2;
						columnOfEachGroup[groupcounter] = (byte)(groupcounter * H_COLUMN_INCREMENT + 1);
						ExcelCellAddress address = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter]);
						groupcounter += 1;
						Group g = new Group(address, new ExcelCellAddress(address.Row + 1, address.Column));
						d.groups.Add((string)sheet.Cells[address.Address].Value, g);
					}
					DefinitionParserData.Item[] entries = GetItemEnties(sheet);
					foreach (DefinitionParserData.Item entry in entries) {
						Group g = d.groups[entry.group];
						g.totalEntries++;
						d.groups[entry.group] = g;

						foreach (KeyValuePair<string, Group> item in d.groups) {
							if ((string)sheet.Cells[item.Value.groupName.Address].Value == entry.group) {
								groupcounter = GetGroupIndex(sheet, (string)sheet.Cells[item.Value.groupName.Address].Value);
							}
						}
						rowOfEachGroup[groupcounter] += 1;
						ExcelCellAddress collected = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter] + 2);
						d.addresses.Add(entry.mainPronounciation, collected);
					}
					return d;
				}
				case SpreadsheetTemplates.SpreadsheetPresetType.ENEMY: {
					ExcelCellAddress current = new ExcelCellAddress(DATA_FIRST_ENTRY);

					while (sheet.Cells[current.Address].Value != null) {
						d.addresses.Add(sheet.Cells[current.Row, current.Column].GetValue<string>(), new ExcelCellAddress(current.Row, current.Column + 2));
						current = SpreadsheetHelper.Advance(sheet, current, out bool nextGroup);
					}
					return d;
				}
				default: {
					throw new CustomException("Uncathegorized sheet entered!");
				}
			}
		}

		private DefinitionParserData.Item[] GetItemEnties(ExcelWorksheet sheet) {
			List<DefinitionParserData.Item> items = new List<DefinitionParserData.Item>();
			byte currColumn = 0;
			ExcelCellAddress currItemAddress = new ExcelCellAddress(DATA_FIRST_ENTRY);
			while (sheet.Cells[currItemAddress.Address].Value != null) {
				string currGroup = (string)sheet.Cells[GROUP_ROW, GROUP_COL + currColumn * H_COLUMN_INCREMENT].Value;

				string mainPron = (string)sheet.GetValue(currItemAddress.Row, currItemAddress.Column);
				uint yangVal = uint.Parse((string)sheet.Cells[SpreadsheetHelper.OffsetAddress(currItemAddress, 0, 1).Address].Value);
				DefinitionParserData.Item item = new DefinitionParserData.Item(
					mainPron,
					new string[0],
					yangVal,
					currGroup
				);
				items.Add(item);
				currItemAddress = SpreadsheetHelper.Advance(sheet, currItemAddress, out bool nextGroup);
				if (nextGroup) {
					currColumn++;
				}
			}
			return items.ToArray();
		}


		public void InitAreaSheet(string areaName, ExcelWorksheet sheet) {
			DefinitionParser parser = new DefinitionParser();
			Dictionaries dicts = InitializeAreaSheet(parser.GetDefinitionByName(areaName), sheet);
			if (!nameToDropCoutAddress.ContainsKey(areaName)) {
				sheetToNameToDropCountAddress.Add(areaName, dicts.addresses);
				sheetToNameToGroupAddress.Add(areaName, dicts.groups);
			}
		}


		public Dictionaries InitializeAreaSheet(DefinitionParserData data, ExcelWorksheet sheet) {
			Dictionaries d = new Dictionaries() {
				addresses = new Dictionary<string, ExcelCellAddress>(),
				groups = new Dictionary<string, Group>()
			};


			sheet.SetValue(SsControl.C_SHEET_NAME, "Spreadsheet for " + sheet.Name);

			int[] rowOfEachGroup = new int[data.groups.Length];
			int[] columnOfEachGroup = new int[data.groups.Length];
			int groupcounter = 0;
			foreach (string group in data.groups) {
				rowOfEachGroup[groupcounter] = GROUP_ROW;
				columnOfEachGroup[groupcounter] = GROUP_COL + groupcounter * H_COLUMN_INCREMENT;
				ExcelCellAddress address = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter]);
				sheet.Select(new ExcelAddress(address.Row, address.Column, address.Row, address.Column + 2));
				ExcelRange r = sheet.SelectedRange;
				r.Merge = true;
				r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				sheet.SetValue(address.Address, group);
				groupcounter += 1;
				Group g = new Group(address, new ExcelCellAddress(address.Row + 1, address.Column));
				d.groups.Add(group, g);
			}
			foreach (DefinitionParserData.Item entry in data.entries) {
				Group g = d.groups[entry.group];
				g.totalEntries++;
				d.groups[entry.group] = g;
				for (int i = 0; i < data.groups.Length; i++) {
					if (data.groups[i] == entry.group) {
						groupcounter = i;
					}
				}
				rowOfEachGroup[groupcounter] += 1;
				ExcelCellAddress nameAddr = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter]);
				ExcelCellAddress yangVal = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter] + 1);
				ExcelCellAddress collected = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter] + 2);

				sheet.SetValue(nameAddr.Address, entry.mainPronounciation);
				sheet.SetValue(yangVal.Address, entry.yangValue);
				sheet.SetValue(collected.Address, 0);

				sheet.Cells[yangVal.Address].Style.Numberformat.Format = "# ###";
				d.addresses.Add(entry.mainPronounciation, collected);
			}
			return d;
		}

		public void InitMobSheet(string mobName, string underlyingArea, ExcelWorksheet sheet) {
			MobAsociatedDrops d = new MobAsociatedDrops();
			Dictionaries dicts = InitializeMobSheet(mobName, underlyingArea, d, sheet);
			if (!sheetToNameToDropCountAddress.ContainsKey(mobName)) {
				sheetToNameToDropCountAddress.Add(mobName, dicts.addresses);
				sheetToNameToGroupAddress.Add(mobName, dicts.groups);
			}
		}

		public Dictionaries InitializeMobSheet(string mobName, string underlyingArea, MobAsociatedDrops data, ExcelWorksheet sheet) {
			Dictionaries d = new Dictionaries() {
				addresses = new Dictionary<string, ExcelCellAddress>(),
				groups = new Dictionary<string, Group>()
			};

			sheet.SetValue(SsControl.C_SHEET_NAME, "Spreadsheet for " + sheet.Name);
			sheet.SetValue(SsControl.E_TOTAL_KILLED, 0);

			Dictionary<string, string[]> itemEntries = data.GetDropsForMob(mobName);

			ExcelCellAddress startAddr = new ExcelCellAddress(GROUP_ROW,GROUP_COL);
			foreach (string key in itemEntries.Keys) {
				sheet.Cells[startAddr.Address].Value = key;
				startAddr = SpreadsheetHelper.OffsetAddress(startAddr, 1, 0);

				for (int i = 0; i < itemEntries[key].Length; i++) {
					ExcelCellAddress itemName = new ExcelCellAddress(startAddr.Row, startAddr.Column);
					ExcelCellAddress yangVal = new ExcelCellAddress(startAddr.Row, startAddr.Column + 1);
					ExcelCellAddress totalDroped = new ExcelCellAddress(startAddr.Row, startAddr.Column + 2);
					sheet.SetValue(itemName.Address, itemEntries[key][i]);
					d.addresses.Add(itemEntries[key][i], totalDroped);
					sheet.SetValue(yangVal.Address, DefinitionParser.instance.GetDefinitionByName(underlyingArea).GetYangValue(itemEntries[key][i]));
					sheet.SetValue(totalDroped.Address, 0);
				}
				startAddr = SpreadsheetHelper.OffsetAddress(startAddr, -1, H_COLUMN_INCREMENT);
			}
			return d;
		}


		private int CountGroups(ExcelWorksheet sheet) {
			ExcelCellAddress start = new ExcelCellAddress(GROUP_ROW, GROUP_COL);
			int counter = 0;
			while (sheet.Cells[start.Address].Value != null && (string)sheet.Cells[start.Address].Value != "GROUP_NAME") {
				counter++;
				start = SpreadsheetHelper.OffsetAddress(start, 0, H_COLUMN_INCREMENT);
			}
			return counter;
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
