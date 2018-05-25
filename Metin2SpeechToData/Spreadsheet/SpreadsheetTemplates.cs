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

		public SpreadsheetTemplates() {	}

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


		/// <summary>
		/// Initializes new Enemy sheet for 'name' in 'current' workbook with given 'data'
		/// </summary>
		public ExcelWorksheet InitEnemySheet(ExcelWorkbook current, string name, MobAsociatedDrops data) {
			ExcelWorksheet sheet = CreateFromTemplate(current, SpreadsheetPresetType.ENEMY, name);
			FillHeadder(current.Worksheets[name], name);
			Dictionary<string, string[]> items = data.GetDropsForMob(name);

			//Group setup
			ExcelCellAddress initGroupAddress = new ExcelCellAddress(GROUP_ROW, GROUP_COL);
			foreach (string group in items.Keys) {
				sheet.SetValue(initGroupAddress.Address, group);
				initGroupAddress = SpreadsheetHelper.OffsetAddress(initGroupAddress.Address, 0, H_COLUMN_INCREMENT);
			}

			//Item setup
			byte currentGroup = 0;
			foreach (string[] parsed in items.Values) {
				for (int i = 0; i < parsed.Length; i++) {
					ExcelCellAddress nameAddr = new ExcelCellAddress(GROUP_ROW, GROUP_COL + currentGroup * H_COLUMN_INCREMENT);

					sheet.SetValue(nameAddr.Row, nameAddr.Column + 0, parsed[i]);
					sheet.SetValue(nameAddr.Row, nameAddr.Column + 1, DefinitionParser.instance.currentGrammarFile.GetYangValue(parsed[i]));
					sheet.SetValue(nameAddr.Row, nameAddr.Column + 2, 0);
					sheet.Cells[nameAddr.Row, nameAddr.Column + 1].Style.Numberformat.Format = "###,###,###,###,###,###,###,###"; // this sucks
				}
				currentGroup += H_COLUMN_INCREMENT;
			}
			return sheet;
		}

		/// <summary>
		/// Initializes Area styled sheet in 'current' workbook according to given item 'data'
		/// </summary>
		public ExcelWorksheet InitAreaSheet(ExcelWorkbook current, DefinitionParserData data) {
			ExcelWorksheet sheet = CreateFromTemplate(current, SpreadsheetPresetType.AREA, data.ID);
			FillHeadder(current.Worksheets[data.ID], data);

			//Group setup
			int[] rowOfEachGroup = new int[data.groups.Length];
			int[] columnOfEachGroup = new int[data.groups.Length];
			int groupcounter = 0;

			ExcelCellAddress initGroupAddress = new ExcelCellAddress(GROUP_ROW, GROUP_COL);
			foreach (string group in data.groups) {
				rowOfEachGroup[groupcounter] = initGroupAddress.Row;
				columnOfEachGroup[groupcounter] = initGroupAddress.Column;
				groupcounter += 1;

				sheet.SetValue(initGroupAddress.Address, group);
				initGroupAddress = SpreadsheetHelper.OffsetAddress(initGroupAddress.Address, 0, H_COLUMN_INCREMENT);
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
			return sheet;
		}

		public ExcelWorksheet InitSessionSheet(ExcelWorkbook current) {
			ExcelWorksheet sheet = CreateFromTemplate(current, SpreadsheetPresetType.SESSION, "Session");
			//TODO some init here
			return sheet;
		}


		/// <summary>
		/// Creates new spreadsheet named 'name' in 'current' workbook of type 'type'
		/// </summary>
		private ExcelWorksheet CreateFromTemplate(ExcelWorkbook current, SpreadsheetPresetType type, string name) {
			ExcelWorksheets _sheets = LoadTemplates();
			switch (type) {
				case SpreadsheetPresetType.MAIN: {
					current.Worksheets.Add(name, _sheets["Main"]);
					break;
				}
				case SpreadsheetPresetType.AREA: {
					current.Worksheets.Add(name, _sheets["Area"]);
					break;
				}
				case SpreadsheetPresetType.ENEMY: {
					current.Worksheets.Add(name, _sheets["Enemy"]);
					break;
				}
				case SpreadsheetPresetType.SESSION: {
					current.Worksheets.Add(name, _sheets["Session"]);
					break;
				}
			}
			_sheets.Dispose();
			Dispose();
			return current.Worksheets[name];
		}

		/// <summary>
		/// Fill headder information of 'curr' according to 'data'
		/// </summary>
		private void FillHeadder(ExcelWorksheet curr, DefinitionParserData data) {
			FormLink(interaction.mainSheet, curr, data.ID);
			curr.SetValue(SsControl.C_SHEET_NAME, data.ID);
			curr.SetValue(SsControl.A_E_LAST_MODIFICATION, DateTime.Now.ToLongDateString());
			curr.SetValue(SsControl.A_E_TOTAL_MERGED_SESSIONS, 0);
		}
		/// <summary>
		/// Fill headder information of 'curr' according to 'data'
		/// </summary>
		private void FillHeadder(ExcelWorksheet curr, string name) {

			FormLink(interaction.mainSheet, curr, name);
			curr.SetValue(SsControl.C_SHEET_NAME, name);
			curr.SetValue(SsControl.A_E_LAST_MODIFICATION, DateTime.Now.ToLongDateString());
			curr.SetValue(SsControl.A_E_TOTAL_MERGED_SESSIONS, 0);
		}

		/// <summary>
		/// Links current sheet with main
		/// </summary>
		private void FormLink(ExcelWorksheet mainSheet, ExcelWorksheet curr, string name) {
			ExcelCellAddress hlinkAddress = new ExcelCellAddress(MAIN_HLINK_OFFSET);
			int currOffsetMain = mainSheet.GetValue<int>(hlinkAddress.Row, hlinkAddress.Column);

			ExcelCellAddress freeLinkSpot = new ExcelCellAddress(MAIN_FIRST_HLINK);
			freeLinkSpot = new ExcelCellAddress(freeLinkSpot.Row + currOffsetMain, freeLinkSpot.Column);

			mainSheet.SetValue(MAIN_HLINK_OFFSET, currOffsetMain + 1);

			SpreadsheetHelper.HyperlinkCell(mainSheet, freeLinkSpot.Address, curr, SsControl.C_RETURN_LINK, name);
			SpreadsheetHelper.HyperlinkCell(curr, SsControl.C_RETURN_LINK, mainSheet, freeLinkSpot.Address, ">>Main Sheet<<");
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
