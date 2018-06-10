using Metin2SpeechToData;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static Metin2SpeechToData.Spreadsheet.SsConstants;


namespace SheetSync {
	class MergeHelper {

		private Dictionary<string, ExcelCellAddress> nameToDropCoutAddress = new Dictionary<string, ExcelCellAddress>();

		private readonly Dictionary<string, Dictionary<string, ExcelCellAddress>> sheetToNameToDropCountAddress = new Dictionary<string, Dictionary<string, ExcelCellAddress>>();

		private readonly List<(string sheetName, SpreadsheetTemplates.SpreadsheetPresetType sheetType)> modifiedLists = new List<(string, SpreadsheetTemplates.SpreadsheetPresetType)>();

		private readonly Dictionary<string, int> totalDropedItemsPerSheet = new Dictionary<string, int>();

		private SessionSheet.ItemMeta[] itemArray;


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
						nameToDropCoutAddress.Clear();
						modifiedLists.Clear();
						totalDropedItemsPerSheet.Clear();
					}
				}
			}
		}


		public void MergeSession(ExcelPackage main, FileInfo fileInfo) {
			ExcelPackage sessionPackage = new ExcelPackage(fileInfo);

			ExcelWorksheet session = sessionPackage.Workbook.Worksheets["Session"];
			session.SetValue(SessionSheet.MERGED_STATUS, "Merged!");

			string currAddress = DATA_FIRST_ENTRY;
			List<SessionSheet.ItemMeta> items = new List<SessionSheet.ItemMeta>();
			while (session.Cells[currAddress].Value != null) {
				DefinitionParserData.Item item = new DefinitionParserData.Item(
					(string)session.Cells[currAddress].Value,
					new string[0],
					uint.Parse(session.Cells[SpreadsheetHelper.OffsetAddressString(currAddress, 0, 13)].Value.ToString()),
					(string)session.Cells[SpreadsheetHelper.OffsetAddressString(currAddress, 0, 4)].Value
				);
				string enemy = (string)session.Cells[SpreadsheetHelper.OffsetAddress(currAddress, 0, 7).Address].Value;
				string currentSheetName = enemy != UNSPEICIFIED_ENEMY ? enemy : (string)session.Cells[SessionSheet.SESSION_AREA_NAME].Value;

				SessionSheet.ItemMeta meta = new SessionSheet.ItemMeta(item, currentSheetName, default(DateTime), 1);
				items.Add(meta);

				if (!modifiedLists.Contains((currentSheetName, (enemy == currentSheetName ? SpreadsheetTemplates.SpreadsheetPresetType.ENEMY : SpreadsheetTemplates.SpreadsheetPresetType.AREA)))) {
					modifiedLists.Add((currentSheetName, (enemy == currentSheetName ? SpreadsheetTemplates.SpreadsheetPresetType.ENEMY : SpreadsheetTemplates.SpreadsheetPresetType.AREA)));
				}
				if (totalDropedItemsPerSheet.ContainsKey(currentSheetName)) {
					totalDropedItemsPerSheet[currentSheetName] = totalDropedItemsPerSheet[currentSheetName] + 1;
				}
				else {
					totalDropedItemsPerSheet.Add(currentSheetName, 1);
				}
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
					nameToDropCoutAddress = sheetToNameToDropCountAddress[currentSheetName];
				}
				else {
					nameToDropCoutAddress = sheetToNameToDropCountAddress[currentSheetName];
				}

				if (nameToDropCoutAddress.ContainsKey(item.mainPronounciation)) {
					int currValue = currMain.GetValue<int>(nameToDropCoutAddress[item.mainPronounciation].Row, nameToDropCoutAddress[item.mainPronounciation].Column);
					currMain.SetValue(nameToDropCoutAddress[item.mainPronounciation].Address, currValue + 1);
				}
				else {
					AddItemEntry(currMain, item);
				}
				currAddress = SpreadsheetHelper.OffsetAddressString(currAddress, 1, 0);
			}
			itemArray = items.ToArray();
			UpdateLinks(main, sessionPackage);
			UpdateSheetHeadders(main);
			main.Save();
			sessionPackage.Save();
		}


		public void UpdateSheetHeadders(ExcelPackage package) {
			for (int i = 0; i < modifiedLists.Count; i++) {
				ExcelWorksheet current = package.Workbook.Worksheets[modifiedLists[i].sheetName];
				int itemCount = CountItems(current);
				current.SetValue(SsControl.C_TOTAL_DROPED_ITEMS, current.Cells[SsControl.C_TOTAL_DROPED_ITEMS].GetValue<int>() + totalDropedItemsPerSheet[modifiedLists[i].sheetName]);
				current.SetValue(SsControl.C_TOTAL_DROPED_VALUE, current.Cells[SsControl.C_TOTAL_DROPED_VALUE].GetValue<int>() + itemArray.Where((x) => x.comesFromEnemy == modifiedLists[i].sheetName).Sum((x) => x.itemBase.yangValue));
				current.SetValue(SsControl.A_E_TOTAL_GROUPS, CountGroups(current));
				current.SetValue(SsControl.C_TOTAL_ITEMS, itemCount);
				current.SetValue(SsControl.A_E_TOTAL_MERGED_SESSIONS, current.Cells[SsControl.A_E_TOTAL_MERGED_SESSIONS].GetValue<int>() + 1);
				current.SetValue(SsControl.A_E_LAST_MODIFICATION, DateTime.Now.ToShortDateString() + " | " + DateTime.Now.ToShortTimeString());

				if (modifiedLists[i].sheetType == SpreadsheetTemplates.SpreadsheetPresetType.ENEMY) {
					current.SetValue(SsControl.E_TOTAL_KILLED, current.Cells[SsControl.E_TOTAL_KILLED].GetValue<int>() + 1);
					current.SetValue(SsControl.E_AVERAGE_DROP, GetAverage(current.Cells[SsControl.E_AVERAGE_DROP].GetValue<float>(), itemCount, itemArray.Where((x) => x.comesFromEnemy == modifiedLists[i].sheetName).ToArray()));
				}
			}
		}

		private int GetAverage(float currAverage, int currItemCount, SessionSheet.ItemMeta[] newItems) {
			Console.WriteLine("Not Implemented");
			return -1;
			throw new NotImplementedException(); //TODO implement average drop calculation
		}

		private int CountItems(ExcelWorksheet current) {
			ExcelCellAddress start = new ExcelCellAddress(DATA_FIRST_ENTRY);
			if (current.Cells[start.Address].Value == null) {
				return 0;
			}
			int counter = 1;
			while (start != null) {
				start = SpreadsheetHelper.Advance(current, start, out bool nextGroup);
				counter++;
			}
			return counter;
		}

		private void UpdateLinks(ExcelPackage main, ExcelPackage session) {
			ExcelWorksheet mainInMain = main.Workbook.Worksheets[H_DEFAULT_SHEET_NAME];
			ExcelWorksheet area = null;
			for (int i = 0; i < modifiedLists.Count; i++) {
				SpreadsheetHelper.HyperlinkAcrossFiles(session.File, "Session", "A1", main.Workbook.Worksheets[modifiedLists[i].sheetName], SsControl.C_LAST_SESSION_LINK, "Last Session");
				if (modifiedLists[i].sheetType == SpreadsheetTemplates.SpreadsheetPresetType.AREA) {
					area = main.Workbook.Worksheets[modifiedLists[i].sheetName];
				}
			}
			if (area == null) {
				throw new CustomException("Session doesn't have primary recognizer attached to it!");
			}

			bool sheetExists = false;

			ExcelCellAddress freeSheetLink = new ExcelCellAddress(MAIN_SHEET_LINKS);
			while (mainInMain.Cells[freeSheetLink.Address].Value != null) {
				if (mainInMain.Cells[freeSheetLink.Address].Value.ToString() == area.Name) {
					sheetExists = true;
				}
				freeSheetLink = SpreadsheetHelper.OffsetAddress(freeSheetLink, 1, 0);
			}
			if (!sheetExists) {
				SpreadsheetHelper.Copy(mainInMain, freeSheetLink.Address, SpreadsheetHelper.OffsetAddress(freeSheetLink, 0, 3).Address,
												   SpreadsheetHelper.OffsetAddress(freeSheetLink, 1, 0).Address, SpreadsheetHelper.OffsetAddress(freeSheetLink, 1, 3).Address);

				SpreadsheetHelper.HyperlinkCell(mainInMain, freeSheetLink.Address, area, "A1", area.Name);
			}
			SpreadsheetHelper.HyperlinkCell(area, SsControl.C_RETURN_LINK, mainInMain, "A1", ">>Main Sheet<<");

			ExcelCellAddress start = new ExcelCellAddress(SsControl.A_ENEMIES_FIRST_LINK);
			byte counter = 0;
			for (int i = 0; i < modifiedLists.Count; i++) {
				if (modifiedLists[i].sheetType != SpreadsheetTemplates.SpreadsheetPresetType.AREA) {
					SpreadsheetHelper.HyperlinkCell(area, start.Address, main.Workbook.Worksheets[modifiedLists[i].sheetName], "A1", modifiedLists[i].sheetName);
					SpreadsheetHelper.HyperlinkCell(main.Workbook.Worksheets[modifiedLists[i].sheetName], SsControl.C_RETURN_LINK, area, "A1", "Back to " + area.Name);
					start = SpreadsheetHelper.OffsetAddress(start, 1, 0);
					counter++;
					if (counter == 5) {
						counter = 0;
						start = SpreadsheetHelper.OffsetAddress(start, -5, 4);
					}
				}
			}

			string sessionName = SpreadsheetHelper.GetSessionName(session.File);
			ExcelCellAddress unmergedStart = new ExcelCellAddress(MAIN_UNMERGED_LINKS);

			while (mainInMain.GetValue<string>(unmergedStart.Row, unmergedStart.Column) != sessionName) {
				unmergedStart = SpreadsheetHelper.OffsetAddress(unmergedStart, 1, 0);
			}
			ExcelCellAddress mergedAddr = FindFreeLinkSpot(mainInMain, new ExcelCellAddress(MAIN_MERGED_LINKS));

			SpreadsheetHelper.Copy(mainInMain, mergedAddr.Address, SpreadsheetHelper.OffsetAddress(mergedAddr, 0, 3).Address,
											   SpreadsheetHelper.OffsetAddress(mergedAddr, 1, 0).Address, SpreadsheetHelper.OffsetAddress(mergedAddr, 1, 3).Address);
			SpreadsheetHelper.HyperlinkAcrossFiles(session.File, "Session", "A1", mainInMain, mergedAddr.Address, sessionName);
			mainInMain.Cells[unmergedStart.Address].Formula = null;
			mainInMain.Cells[unmergedStart.Address].Value = null;
			mainInMain.Select("A1");
		}


		private ExcelCellAddress FindFreeLinkSpot(ExcelWorksheet sheet, ExcelCellAddress start) {
			while (sheet.GetValue(start.Row, start.Column) != null) {
				start = SpreadsheetHelper.OffsetAddress(start, 1, 0);
			}
			return start;
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
						rowOfEachGroup[groupcounter] = GROUP_ROW;
						columnOfEachGroup[groupcounter] = (byte)(GROUP_COL + groupcounter * H_COLUMN_INCREMENT);
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

					while (current != null && sheet.Cells[current.Address].Value != null) {
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
			while (currItemAddress != null && sheet.Cells[currItemAddress.Address].Value != null) {
				string currGroup = (string)sheet.Cells[GROUP_ROW, GROUP_COL + currColumn * H_COLUMN_INCREMENT].Value;

				string mainPron = (string)sheet.GetValue(currItemAddress.Row, currItemAddress.Column);
				uint yangVal = sheet.GetValue<uint>(currItemAddress.Row, currItemAddress.Column + 1);
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

			Dictionaries dicts = InitializeAreaSheet(DefinitionParser.instance.GetDefinitionByName(areaName), sheet);
			if (!nameToDropCoutAddress.ContainsKey(areaName)) {
				sheetToNameToDropCountAddress.Add(areaName, dicts.addresses);
			}
		}


		public Dictionaries InitializeAreaSheet(DefinitionParserData data, ExcelWorksheet sheet) {
			Dictionaries d = new Dictionaries() {
				addresses = new Dictionary<string, ExcelCellAddress>(),
				groups = new Dictionary<string, Group>()
			};

			sheet.Cells[SsControl.C_SHEET_NAME].Value = "Spreadsheet for " + sheet.Name;

			int[] rowOfEachGroup = new int[data.groups.Length];
			int[] columnOfEachGroup = new int[data.groups.Length];
			int groupCounter = 0;

			foreach (string group in data.groups) {
				rowOfEachGroup[groupCounter] = GROUP_ROW;
				columnOfEachGroup[groupCounter] = GROUP_COL + groupCounter * H_COLUMN_INCREMENT;
				ExcelCellAddress address = new ExcelCellAddress(rowOfEachGroup[groupCounter], columnOfEachGroup[groupCounter]);
				sheet.SetValue(address.Address, group);
				groupCounter += 1;
				Group g = new Group(address, new ExcelCellAddress(address.Row + 1, address.Column));
				d.groups.Add(group, g);
			}
			foreach (DefinitionParserData.Item entry in data.entries) {
				Group g = d.groups[entry.group];
				g.totalEntries++;
				d.groups[entry.group] = g;
				for (int i = 0; i < data.groups.Length; i++) {
					if (data.groups[i] == entry.group) {
						groupCounter = i;
					}
				}
				rowOfEachGroup[groupCounter] += 1;
				ExcelCellAddress nameAddr = new ExcelCellAddress(rowOfEachGroup[groupCounter], columnOfEachGroup[groupCounter]);
				ExcelCellAddress yangVal = new ExcelCellAddress(rowOfEachGroup[groupCounter], columnOfEachGroup[groupCounter] + 1);
				ExcelCellAddress collected = new ExcelCellAddress(rowOfEachGroup[groupCounter], columnOfEachGroup[groupCounter] + 2);

				sheet.SetValue(nameAddr.Address, entry.mainPronounciation);
				sheet.SetValue(yangVal.Address, entry.yangValue);
				sheet.SetValue(collected.Address, 0);
				d.addresses.Add(entry.mainPronounciation, collected);
			}
			return d;
		}

		public void InitMobSheet(string mobName, string underlyingArea, ExcelWorksheet sheet) {
			Dictionaries dicts = InitializeMobSheet(mobName, underlyingArea, sheet);
			if (!sheetToNameToDropCountAddress.ContainsKey(mobName)) {
				sheetToNameToDropCountAddress.Add(mobName, dicts.addresses);
			}
		}

		public Dictionaries InitializeMobSheet(string mobName, string underlyingArea, ExcelWorksheet sheet) {
			Dictionaries d = new Dictionaries() {
				addresses = new Dictionary<string, ExcelCellAddress>(),
				groups = new Dictionary<string, Group>()
			};

			sheet.SetValue(SsControl.C_SHEET_NAME, "Spreadsheet for " + sheet.Name);
			sheet.SetValue(SsControl.E_TOTAL_KILLED, 0);
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
