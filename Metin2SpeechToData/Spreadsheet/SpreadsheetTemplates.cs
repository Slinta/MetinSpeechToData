using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Metin2SpeechToData.Structures;
using System;
using static Metin2SpeechToData.Spreadsheet.SsConstants;

namespace Metin2SpeechToData {
	public class SpreadsheetTemplates : IDisposable {

		public enum SpreadsheetPresetType {
			MAIN,
			AREA,
			ENEMY,
			SESSION
		}

		private readonly SpreadsheetInteraction interaction;
		private ExcelPackage package;

		public SpreadsheetTemplates(SpreadsheetInteraction interaction) {
			this.interaction = interaction;
		}

		/// <summary>
		/// Opens and load Template.xlsx file from Templates folder
		/// <para>Used for copying the template, after the procedure call Dispose!</para>
		/// </summary>
		/// <returns></returns>
		public ExcelWorksheets LoadTemplates() {
			FileInfo templatesFile = new FileInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar +
												  "Templates" + Path.DirectorySeparatorChar + "Templates.xlsx");

			if (!templatesFile.Exists) {
				throw new CustomException("Unable to locate 'Templates.xlsx' inside Templates folder. Redownload might be necessary");
			}
			package = new ExcelPackage(templatesFile);
			return package.Workbook.Worksheets;
		}

		[Obsolete("Rewritten as \"InitMobSheet()\", this function will be removed once the shift to spreadshhet templates is complete.")]
		public Dicts InitializeMobSheet(string mobName, MobAsociatedDrops data) {

			Dicts d = new Dicts(true);

			interaction.InsertValue(new ExcelCellAddress(1, 1), "Spreadsheet for enemy: " + interaction.currentSheet.Name);
			interaction.InsertValue(new ExcelCellAddress(1, 4), "Num killed:");
			interaction.InsertValue(new ExcelCellAddress(1, 5), 0);

			Dictionary<string, string[]> itemEntries = Program.gameRecognizer.enemyHandling.mobDrops.GetDropsForMob(mobName);

			ExcelCellAddress startAddr = new ExcelCellAddress("A2");
			foreach (string key in itemEntries.Keys) {
				interaction.currentSheet.Cells[startAddr.Address].Value = key;

				for (int i = 0; i < itemEntries[key].Length; i++) {
					int _offset = i + 1;
					ExcelCellAddress itemName = new ExcelCellAddress(startAddr.Row + _offset, startAddr.Column);
					ExcelCellAddress yangVal = new ExcelCellAddress(startAddr.Row + _offset, startAddr.Column + 1);
					ExcelCellAddress totalDroped = new ExcelCellAddress(startAddr.Row + _offset, startAddr.Column + 2);
					interaction.InsertValue(itemName, itemEntries[key][i]);
					d.addresses.Add(itemEntries[key][i], totalDroped);
					interaction.InsertValue(yangVal, DefinitionParser.instance.currentGrammarFile.GetYangValue(itemEntries[key][i]));
					interaction.InsertValue(totalDroped, 0);
				}
				startAddr = new ExcelCellAddress(2, startAddr.Column + 4);
			}
			interaction.Save();
			return d;
		}

		public void InitEnemySheet(ExcelWorkbook current, string name, MobAsociatedDrops data) {
			//TODO add template for enemy sheet
			//TODO complete this function
		}

		[Obsolete("Rewritten as \"InitAreaSheet()\", this function will be removed once the shift to spreadshhet templates is complete.")]
		public Dicts InitializeAreaSheet(DefinitionParserData data) {
			Dicts d = new Dicts(true);

			interaction.InsertValue(new ExcelCellAddress(1, 1), "Spreadsheet for zone: " + interaction.currentSheet.Name);

			int[] rowOfEachGroup = new int[data.groups.Length];
			int[] columnOfEachGroup = new int[data.groups.Length];
			int groupcounter = 0;
			foreach (string group in data.groups) {
				rowOfEachGroup[groupcounter] = 2;
				columnOfEachGroup[groupcounter] = groupcounter * H_COLUMN_INCREMENT + 1;
				ExcelCellAddress address = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter]);
				interaction.currentSheet.Select(new ExcelAddress(address.Row, address.Column, address.Row, address.Column + 2));
				ExcelRange r = interaction.currentSheet.SelectedRange;
				r.Merge = true;
				r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				interaction.InsertValue(address, group);
				groupcounter += 1;
				SpreadsheetInteraction.Group g = new SpreadsheetInteraction.Group(address, new ExcelCellAddress(address.Row + 1, address.Column));
				d.groups.Add(group, g);
			}
			foreach (DefinitionParserData.Item entry in data.entries) {
				SpreadsheetInteraction.Group g = d.groups[entry.group];
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

				interaction.InsertValue(nameAddr, entry.mainPronounciation);
				interaction.InsertValue(yangVal, entry.yangValue);
				interaction.InsertValue(collected, 0);

				interaction.currentSheet.Cells[yangVal.Address].Style.Numberformat.Format = "###,###,###,###,###";
				d.addresses.Add(entry.mainPronounciation, collected);
			}
			interaction.Save();
			return d;
		}

		/// <summary>
		/// Initializes Area styled sheet in 'current' workbook according to given item 'data'
		/// </summary>
		/// <param name="current"></param>
		/// <param name="data"></param>
		public void InitAreaSheet(ExcelWorkbook current, DefinitionParserData data) {
			ExcelWorksheet sheet = CreateFromTemplate(current, SpreadsheetPresetType.AREA, data.ID);
			FillHeadder(current.Worksheets[data.ID], data, SpreadsheetPresetType.AREA);

			//Group setup
			int[] rowOfEachGroup = new int[data.groups.Length];
			int[] columnOfEachGroup = new int[data.groups.Length];
			int groupcounter = 0;

			ExcelCellAddress initGroupAddress = new ExcelCellAddress(D_GROUP);
			foreach (string group in data.groups) {
				rowOfEachGroup[groupcounter] = initGroupAddress.Row;
				columnOfEachGroup[groupcounter] = initGroupAddress.Column;
				groupcounter += 1;

				sheet.SetValue(initGroupAddress.Address, group);
				initGroupAddress = new ExcelCellAddress(SpreadsheetHelper.OffsetAddress(initGroupAddress.Address, 0, H_COLUMN_INCREMENT));
			}

			//Item setup
			foreach (DefinitionParserData.Item entry in data.entries) {
				for (int i = 0; i < data.groups.Length; i++) {
					if (data.groups[i] == entry.group) {
						groupcounter = i;
					}
				}
				rowOfEachGroup[groupcounter] += 1;
				ExcelCellAddress nameAddr = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter]);

				sheet.SetValue(nameAddr.Row, nameAddr.Column + 0, entry.mainPronounciation);
				sheet.SetValue(nameAddr.Row, nameAddr.Column + 1, entry.yangValue);
				sheet.SetValue(nameAddr.Row, nameAddr.Column + 2, 0);
				sheet.Cells[nameAddr.Row, nameAddr.Column + 1].Style.Numberformat.Format = "###,###,###,###,###,###,###,###"; // this sucks
			}
		}

		/// <summary>
		/// Creates new spreadsheet named 'name' in 'current' workbook of type 'type'
		/// </summary>
		private ExcelWorksheet CreateFromTemplate(ExcelWorkbook current, SpreadsheetPresetType type, string name) {
			ExcelWorksheets _sheets = LoadTemplates();
			switch (type) {
				case SpreadsheetPresetType.MAIN: {
					interaction.currentSheet.Workbook.Worksheets.Add(name, _sheets["Main"]);
					break;
				}
				case SpreadsheetPresetType.AREA: {
					interaction.currentSheet.Workbook.Worksheets.Add(name, _sheets["Area"]);
					break;
				}
				case SpreadsheetPresetType.ENEMY: {
					interaction.currentSheet.Workbook.Worksheets.Add(name, _sheets["Enemy"]);
					break;
				}
				case SpreadsheetPresetType.SESSION: {
					interaction.currentSheet.Workbook.Worksheets.Add(name, _sheets["Session"]);
					break;
				}
			}
			interaction.currentSheet.Workbook.Worksheets.Add(name, _sheets["Area"]);
			_sheets.Dispose();
			Dispose();
			return current.Worksheets[name];
		}

		/// <summary>
		/// Fill headder information of 'curr' according to 'data' and 'type'
		/// </summary>
		private void FillHeadder(ExcelWorksheet curr, DefinitionParserData data, SpreadsheetPresetType type) {
			FormLink(interaction.mainSheet, curr, data.ID);
			if (type == SpreadsheetPresetType.AREA) {
				curr.SetValue(SsControl.C_SHEET_NAME, data.ID);
				curr.SetValue(SsControl.C_LAST_MODIFICATION, DateTime.Now);
				curr.SetValue(SsControl.C_TOTAL_MERGED_SESSIONS, 0);
			}
		}

		/// <summary>
		/// Links current sheet with main
		/// </summary>
		private void FormLink(ExcelWorksheet mainSheet, ExcelWorksheet curr, string name) {
			ExcelCellAddress hlinkAddress = new ExcelCellAddress(C_HLINK_OFFSET);
			int currOffsetMain = mainSheet.GetValue<int>(hlinkAddress.Row, hlinkAddress.Column);

			ExcelCellAddress freeLinkSpot = new ExcelCellAddress(C_FIRST_HLINK);
			freeLinkSpot = new ExcelCellAddress(freeLinkSpot.Row + currOffsetMain, freeLinkSpot.Column);

			mainSheet.SetValue(C_HLINK_OFFSET, currOffsetMain + 1);

			SpreadsheetHelper.HyperlinkCell(mainSheet, freeLinkSpot.Address, curr, SsControl.C_RETURN_LINK, name);
			SpreadsheetHelper.HyperlinkCell(curr, SsControl.C_RETURN_LINK, mainSheet, freeLinkSpot.Address, ">Main Sheet<");
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				try {
					package.Dispose();
				}
				catch {
					//Attempt to dispose
				}
				disposedValue = true;
			}
		}

		~SpreadsheetTemplates() {
			Dispose(false);
		}

		// This code added to correctly implement the disposable pattern.
		public void Dispose() {
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
