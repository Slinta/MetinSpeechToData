using Metin2SpeechToData;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using static Metin2SpeechToData.Spreadsheet.SsConstants;


namespace SheetSync {
	class MergeHelper {

		private static FileInfo[] sessions;
		private Dictionary<string, string> nameToDropCoutAddress = new Dictionary<string, string>();
		private Dictionary<string, Dictionary<string, string>> sheetToNameToDropCountAddress = new Dictionary<string, Dictionary<string, string>>();



		private static void MergeSession(ExcelPackage main, FileInfo fileInfo) {
			//ExcelPackage package = new ExcelPackage(fileInfo);
			//ExcelWorksheet session = package.Workbook.Worksheets["Session"];

			//string currAddress = DATA_FIRST_ENTRY;

			//ExcelWorksheet currMain = main.Workbook.Worksheets[H_DEFAULT_SHEET_NAME];
			//while (session.Cells[currAddress].Value != null) {
			//	DefinitionParserData.Item item = new DefinitionParserData.Item((string)session.Cells[currAddress].Value,
			//		new string[0],
			//		uint.Parse(session.Cells[SpreadsheetHelper.OffsetAddressString(currAddress, 0, 13)].Value.ToString()),
			//		(string)session.Cells[SpreadsheetHelper.OffsetAddressString(currAddress, 0, 4)].Value);
			//	string enemy = (string)session.Cells[SpreadsheetHelper.OffsetAddress(currAddress, 0, 7).Address].Value;


			//	if (enemy != UNSPEICIFIED_ENEMY) {
			//		try {
			//			if (currMain != main.Workbook.Worksheets[enemy]) {
			//				nameToDropCoutAddress = sheetToNameToDropCountAddress[enemy];
			//			}
			//			currMain = main.Workbook.Worksheets[enemy];
			//		}
			//		catch {
			//			SpreadsheetTemplates t = new SpreadsheetTemplates();
			//			ExcelWorksheets sheets = t.LoadTemplates();
			//			main.Workbook.Worksheets.Add(enemy, sheets["Enemy"]);
			//			t.Dispose();

			//			if (currMain != main.Workbook.Worksheets[enemy]) {
			//				if (!sheetToNameToDropCountAddress.ContainsKey(enemy)) {
			//					sheetToNameToDropCountAddress.Add(enemy, new Dictionary<string, string>());
			//				}
			//				nameToDropCoutAddress = sheetToNameToDropCountAddress[enemy];
			//			}
			//			currMain = main.Workbook.Worksheets[enemy];
			//		}
			//		if (nameToDropCoutAddress.ContainsKey(item.mainPronounciation)) {
			//			currMain.SetValue(nameToDropCoutAddress[item.mainPronounciation], (int)currMain.Cells[nameToDropCoutAddress[item.mainPronounciation]].Value + 1);
			//		}
			//		else {
			//			AddItemEntry(currMain, item);
			//		}
			//	}
			//	else {
			//		try {
			//			string areaName = (string)session.Cells[SessionSheet.SESSION_AREA_NAME].Value;
			//			if (currMain != main.Workbook.Worksheets[areaName]) {
			//				if (!sheetToNameToDropCountAddress.ContainsKey(areaName)) {
			//					sheetToNameToDropCountAddress.Add(main.Workbook.Worksheets[areaName].Name, new Dictionary<string, string>());
			//				}
			//				nameToDropCoutAddress = sheetToNameToDropCountAddress[areaName];
			//			}
			//			currMain = main.Workbook.Worksheets[areaName];
			//		}
			//		catch {
			//			string areaName = (string)session.Cells[SessionSheet.SESSION_AREA_NAME].Value;
			//			SpreadsheetTemplates t = new SpreadsheetTemplates();
			//			ExcelWorksheets sheets = t.LoadTemplates();
			//			main.Workbook.Worksheets.Add(areaName, sheets["Area"]);

			//			if (currMain != main.Workbook.Worksheets[areaName]) {
			//				if (!sheetToNameToDropCountAddress.ContainsKey(areaName)) {
			//					sheetToNameToDropCountAddress.Add(areaName, new Dictionary<string, string>());
			//				}
			//				nameToDropCoutAddress = sheetToNameToDropCountAddress[areaName];
			//			}
			//			currMain = main.Workbook.Worksheets[areaName];
			//		}


			//		if (nameToDropCoutAddress.ContainsKey(item.mainPronounciation)) {
			//			currMain.SetValue(nameToDropCoutAddress[item.mainPronounciation], (int)currMain.Cells[nameToDropCoutAddress[item.mainPronounciation]].Value + 1);
			//		}
			//		else {
			//			AddItemEntry(currMain, item);
			//		}
			//	}
			//	currAddress = SpreadsheetHelper.OffsetAddressString(currAddress, 1, 0);
			//}
			//main.Save();
		}


		private static void AddItemEntry(ExcelWorksheet currentSheet, DefinitionParserData.Item entry) {
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
			nameToDropCoutAddress.Add(entry.mainPronounciation, new ExcelCellAddress(current.Row, current.Column + 2).Address);
		}

		private static short GetGroupIndex(ExcelWorksheet sheet, string group) {
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
		public static Dicts LoadSpreadsheet(ExcelWorksheet sheet, SpreadsheetTemplates.SpreadsheetPresetType type) {
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
						current = SpreadsheetHelper.Advance(sheet, current);
					}
					return d;
				}
				default: {
					throw new CustomException("Uncathegorized sheet entered!");
				}
			}
		}


	}
}
