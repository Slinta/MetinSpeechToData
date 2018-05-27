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
		}
		#endregion

		public void StartSession(string grammar) {
			if (currentSession != null) {
				currentSession.Finish();
				currentSession = new SessionSheet(this, grammar, xlsxFile.File);
			}
			else {
				currentSession = new SessionSheet(this, grammar, xlsxFile.File);
			}
		}

		public void StopSession() {
			currentSession.Finish();
			currentSession = null;
		}

		public string UnmergedLinkSpot(string sessionName) {
			string currentAddr = MAIN_UNMERGED_LINKS;

			while (currentSheet.Cells[currentAddr].Value != null) {
				currentAddr = SpreadsheetHelper.OffsetAddress(currentAddr, 1, 0).Address;
			}
			currentSheet.Select(currentAddr + ":" + SpreadsheetHelper.OffsetAddressString(currentAddr,0,3));
			ExcelRange yellow = currentSheet.SelectedRange;
			currentSheet.Select(SpreadsheetHelper.OffsetAddress(currentAddr, 1, 0).Address + ":" + SpreadsheetHelper.OffsetAddress(currentAddr, 1, 3).Address);
			yellow.Copy(currentSheet.SelectedRange);

			SpreadsheetHelper.HyperlinkAcrossFiles(currentSession.package.File, "Session", SessionSheet.LINK_TO_MAIN, mainSheet, currentAddr, sessionName);
			currentSheet.Select("A1");
			Save();
			return currentAddr;
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
