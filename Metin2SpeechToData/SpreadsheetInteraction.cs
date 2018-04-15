using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public class SpreadsheetInteraction {
		/// <summary>
		/// Excel file
		/// </summary>
		private ExcelPackage xlspackage;
		/// <summary>
		/// Excel workbook
		/// </summary>
		private ExcelWorkbook xlsworkbook;
		/// <summary>
		/// Excel sheet
		/// </summary>
		private ExcelWorksheet xlssheet;

		/// <summary>
		/// Open spreadsheet
		/// </summary>
		public SpreadsheetInteraction(string path, string worksheetName) {
			xlspackage = new ExcelPackage(new FileInfo(path));
			xlsworkbook = xlspackage.Workbook;
			OpenWorksheet(worksheetName);
			xlspackage.Save();
		}
		/// <summary>
		/// Open excel file
		/// </summary>
		public SpreadsheetInteraction(string path) {
			xlspackage = new ExcelPackage(new FileInfo(path));
			xlsworkbook = xlspackage.Workbook;
		}

		/// <summary>
		/// Open existing worksheet or add one if it does't exist
		/// </summary>
		public void OpenWorksheet(string sheetname) {
			if (xlsworkbook.Worksheets[sheetname] == null) {
				xlsworkbook.Worksheets.Add(sheetname);
			}
			xlssheet = xlsworkbook.Worksheets[sheetname];
		}

		/// <summary>
		/// Add a sheet by index
		/// </summary>
		public void OpenWorksheet(int sheetindex) {
			if (xlsworkbook.Worksheets[sheetindex] == null) {
				throw new Exception("No worksheet with id " + sheetindex + " exists");
			}
			else {
				xlssheet = xlsworkbook.Worksheets[sheetindex];
			}
		}

		/// <summary>
		/// Adds a specified number to the number in a single cell specified by the address and saves the document
		/// </summary>
		public void AddNumberTo (ExcelCellAddress address, int number) {
			//If there's no sheet then you can't edit it
			if (xlssheet == null) {
				Console.WriteLine("No sheet set");
				return;
			}
			//The value in the cell
			int numberInCell;
			//If the cell is empty change the value anyway
			if(xlssheet.Cells[address.Row, address.Column].Value == null) {
				xlssheet.Cells[address.Row, address.Column].Value = number;
				Console.WriteLine("Changed cell [" + address.Row + " " + address.Column + "] to " + number + ", cell was empty");

				xlspackage.Save();
				return;
			}
			//if the cell already has a number add the new one to the existing one
			if (int.TryParse(xlssheet.Cells[address.Row, address.Column].Value.ToString(), out numberInCell)) {
				numberInCell += number;
				xlssheet.Cells[address.Row, address.Column].Value = numberInCell;
				Console.WriteLine("Changed cell [" + address.Row + " " + address.Column + "] to " + number + ", cell was empty");

				xlspackage.Save();
				return;		
			}
			//if no criterium is met write the error
			Console.WriteLine("Failed to add " + number + " to cell[" + address.Row + " " + address.Column + "], cell doesn't contain a number");
		}

		public void InsertText(ExcelCellAddress address, string text) {
			//If there's no sheet then you can't edit it
			if (xlssheet == null) {
				Console.WriteLine("No sheet set");
				return;
			}
			
			//change the cell

			xlssheet.Cells[address.Row, address.Column].Value = text;
			Console.WriteLine("Changed cell [" + address.Row + " " + address.Column + "] to " + text + ", cell was empty");
			xlspackage.Save();
			return;
		}
	}
}
