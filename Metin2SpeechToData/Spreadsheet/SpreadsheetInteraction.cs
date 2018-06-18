using System;
using System.IO;
using OfficeOpenXml;
using System.Linq;

using static Metin2SpeechToData.Configuration;
using static Metin2SpeechToData.Spreadsheet.SsConstants;

namespace Metin2SpeechToData {
	public class SpreadsheetInteraction : IDisposable {

		private readonly ExcelPackage xlsxFile;
		private readonly ExcelWorkbook content;

		public ExcelWorksheet mainSheet { get; }

		private uint currModificationsToXlsx = 0;

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
			mainSheet = content.Worksheets["Metin2 Drop Analyzer"];
		}


		#endregion

		/// <summary>
		/// Starts session with selected 'grammar'
		/// </summary>
		public void StartSession(string grammar) {
			if (currentSession != null) {
				currentSession.Finish();
				currentSession = new SessionSheet(this, grammar, xlsxFile.File);
			}
			else {
				currentSession = new SessionSheet(this, grammar, xlsxFile.File);
			}
		}

		/// <summary>
		/// Finalizes session
		/// </summary>
		public void StopSession() {
			if (currentSession != null) {
				currentSession.Finish();
				currentSession = null;
			}
		}
		/// <summary>
		/// Helper function for making a link to mainSheet to correct cell
		/// </summary>
		public string UnmergedLinkSpot(string sessionName) {
			string currentAddr = MAIN_UNMERGED_LINKS;

			while (mainSheet.Cells[currentAddr].Value != null) {
				currentAddr = SpreadsheetHelper.OffsetAddress(currentAddr, 1, 0).Address;
			}
			SpreadsheetHelper.Copy(mainSheet, currentAddr, SpreadsheetHelper.OffsetAddress(currentAddr, 0, 3).Address,
											  SpreadsheetHelper.OffsetAddress(currentAddr, 1, 0).Address, SpreadsheetHelper.OffsetAddress(currentAddr, 1, 3).Address);

			SpreadsheetHelper.HyperlinkAcrossFiles(currentSession.package.File, "Session", SessionSheet.LINK_TO_MAIN, mainSheet, currentAddr, sessionName);
			mainSheet.Cells[currentAddr].Calculate();
			Save();
			return currentAddr;
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
				mainSheet.Dispose();
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
