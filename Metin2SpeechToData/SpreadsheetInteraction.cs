using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public class SpreadsheetInteraction {
		/// <summary>
		/// Excel file
		/// </summary>
		ExcelPackage xlspackage;
		/// <summary>
		/// Excel workbook
		/// </summary>
		ExcelWorkbook xlsworkbook;
		/// <summary>
		/// Excel sheet
		/// </summary>
		ExcelWorksheet xlssheet;

		/// <summary>
		/// Open spreadsheet
		/// </summary>
		/// <param name="path"></param>
		/// <param name="worksheetName"></param>
		public SpreadsheetInteraction(string path, string worksheetName) {
			xlspackage = new ExcelPackage(new FileInfo(path));
			xlsworkbook = xlspackage.Workbook;
			xlssheet = xlsworkbook.Worksheets[worksheetName];

		}
		/// <summary>
		/// Open excel file
		/// </summary>
		/// <param name="path"></param>
		public SpreadsheetInteraction(string path) {
			xlspackage = new ExcelPackage(new FileInfo(path));
			xlsworkbook = xlspackage.Workbook;
		}

		/// <summary>
		/// Add a sheet by name
		/// </summary>
		/// <param name="sheetname"></param>
		public void OpenWorksheet(string sheetname) {
			xlssheet = xlsworkbook.Worksheets[sheetname];
		}
		/// <summary>
		/// Add a sheet by index
		/// </summary>
		/// <param name="sheetindex"></param>
		public void OpenWorksheet(int sheetindex) {
			xlssheet = xlsworkbook.Worksheets[sheetindex];
		}

		/// <summary>
		/// Adds a specified number to the number in a single cell specified by the address and saves the document
		/// </summary>
		/// <param name="address"></param>
		/// <param name="number"></param>
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

	}
}
