using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
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

		public SpreadsheetTemplates() { }

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


		public ExcelWorksheet InitSessionSheet(ExcelWorkbook current) {
			ExcelWorksheet sheet = CreateFromTemplate(current, SpreadsheetPresetType.SESSION, "Session");
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
