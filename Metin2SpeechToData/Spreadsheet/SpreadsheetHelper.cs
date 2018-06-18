using OfficeOpenXml;
using System;
using System.IO;
using static Metin2SpeechToData.Spreadsheet.SsConstants;

namespace Metin2SpeechToData {
	public class SpreadsheetHelper {

		/// <summary>
		/// If given address containing itemName
		/// this function will keep returning following addresses,
		/// until it reaches an empty cell in next data column, then it returns null
		/// <para>A3 > A4 > A5 .. Ax is null continue next column E3 > E4 ... I3 is empty >> NULL</para>
		/// </summary>
		public static ExcelCellAddress Advance(ExcelWorksheet sheet, ExcelCellAddress currAddr, out bool nextGroup) {
			currAddr = new ExcelCellAddress(currAddr.Row + 1, currAddr.Column);
			if (sheet.Cells[currAddr.Address].Value == null) {
				currAddr = new ExcelCellAddress(ITEM_ROW, currAddr.Column + H_COLUMN_INCREMENT);
				nextGroup = true;
				if (sheet.Cells[currAddr.Address].Value == null) {
					return null;
				}
			}
			nextGroup = false;
			return currAddr;
		}

		/// <summary>
		/// Creates a hyperlink in 'currentSheet' at 'currentCellAddress' pointing to 'otherSheet' to 'locationAddress', hide link syntax with 'displeyText'
		/// </summary>
		public static void HyperlinkCell(ExcelWorksheet currentSheet, string currentCellAddress, ExcelWorksheet otherSheet, string locationAddress, string displayText) {
			if (Configuration.sheetViewer == Configuration.SheetViewer.EXCEL) {
				currentSheet.Cells[currentCellAddress].Formula =
					string.Format("HYPERLINK(\"#'{0}'!{1}\",\"{2}\")", otherSheet.Name, locationAddress, displayText);
			}
			else {
				currentSheet.Cells[currentCellAddress].Formula =
					string.Format("HYPERLINK(\"#{0}.{1}\",\"{2}\")", otherSheet.Name, locationAddress, displayText);
			}
			currentSheet.Cells[currentCellAddress].Calculate();
		}

		/// <summary>
		/// Copies a range 'fromA' - 'toA' to 'fromB' - 'toB' in 'sheet'
		/// </summary>
		public static void Copy(ExcelWorksheet sheet, string fromStart, string fromEnd, string toStart, string toEnd) {
			sheet.Select(fromStart + ":" + fromEnd);
			ExcelRange from = sheet.SelectedRange;
			sheet.Select(toStart + ":" + toEnd);
			ExcelRange to = sheet.SelectedRange;

			try {
				from.Copy(to);
			}
			catch {

				string c = from.Style.Fill.BackgroundColor.Rgb;
				int[] argb = new int[4];
				int index = 0;
				for (int i = 0; i < c.Length; i += 2) {
					argb[index] = Convert.ToInt32(c.Substring(i, 2), 16);
					index++;
				}

				to.Value = from.Value;
				to.Formula = from.Formula;
				to.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				to.Style.Fill.BackgroundColor.SetColor(argb[0], argb[1], argb[2], argb[3]);
				to.Calculate();
			}
			if (from.Merge) {
				try {
					to.Merge = true;
				}
				catch {
					Console.WriteLine("Unable to merge address {0}.", to.ToString());
					System.Collections.Generic.IEnumerator<ExcelRangeBase> enm = to.GetEnumerator();
					while (enm.MoveNext()) {
						Console.WriteLine("Cell {0} is Merged={1}", enm.Current, enm.Current.Merge);
					}
				}
			}
		}

		/// <summary>
		/// Cuts a range 'fromA' - 'toA' to 'fromB' - 'toB' in 'sheet'
		/// </summary>
		public static void Cut(ExcelWorksheet sheet, string fromStart, string fromEnd, string toStart, string toEnd) {
			Copy(sheet, fromStart, fromEnd, toStart, toEnd);

			sheet.Select(toStart + ":" + toEnd);
			sheet.SelectedRange.Value = null;
		}


		/// <summary>
		/// Creates a hyperlink in 'currentSheet' at 'currentCellAddress' pointing to 'other' file containing 'otherSheet' to defined 'locationAddress', hide link syntax with 'displeyText'
		/// </summary>
		public static void HyperlinkAcrossFiles(FileInfo other, string otherSheet, string otherAddress, ExcelWorksheet current, string currAddress, string displayText) {
			if (Configuration.sheetViewer == Configuration.SheetViewer.EXCEL) {
				current.Cells[currAddress].Formula = string.Format("HYPERLINK(\"[{0}]'{1}'!{2}\",\"{3}\")", other.FullName, otherSheet, otherAddress, displayText);
			}
			else {
				current.Cells[currAddress].Formula = string.Format("HYPERLINK(\"file:///{0}#{1}.{2}\",\"{3}\")", other.FullName, otherSheet, otherAddress, displayText);
			}
			current.Cells[currAddress].Calculate();
		}

		/// <summary>
		/// Gets session name of selected file (as specified in cell B2)
		/// </summary>
		public static string GetSessionName(FileInfo session) {
			ExcelCellAddress nameAddr = new ExcelCellAddress(SsControl.C_SHEET_NAME);
			using (ExcelPackage p = new ExcelPackage(session)) {
				string name = p.Workbook.Worksheets[1].GetValue<string>(nameAddr.Row, nameAddr.Column);
				name = name.Split(':', '|')[1];
				return name.Trim();
			}
		}

		/// <summary>
		/// Takes 'current' address and offsets it by 'rowOffset' rows and'colOffset' columns 
		/// </summary>
		public static string OffsetAddressString(string current, int rowOffset, int colOffset) {
			ExcelCellAddress a = new ExcelCellAddress(current);
			return new ExcelCellAddress(a.Row + rowOffset, a.Column + colOffset).Address;
		}

		/// <summary>
		/// Takes 'current' address and offsets it by 'rowOffset' rows and'colOffset' columns 
		/// </summary>
		public static string OffsetAddressString(ExcelCellAddress current, int rowOffset, int colOffset) {
			return new ExcelCellAddress(current.Row + rowOffset, current.Column + colOffset).Address;
		}

		/// <summary>
		/// Takes 'current' address and offsets it by 'rowOffset' rows and'colOffset' columns 
		/// </summary>
		public static ExcelCellAddress OffsetAddress(ExcelCellAddress current, int rowOffset, int colOffset) {
			return new ExcelCellAddress(current.Row + rowOffset, current.Column + colOffset);
		}

		/// <summary>
		/// Takes 'current' address and offsets it by 'rowOffset' rows and'colOffset' columns 
		/// </summary>
		public static ExcelCellAddress OffsetAddress(string current, int rowOffset, int colOffset) {
			ExcelCellAddress a = new ExcelCellAddress(current);
			return new ExcelCellAddress(a.Row + rowOffset, a.Column + colOffset);
		}

		#region Formula functions
		/// <summary>
		/// Sums cells in 'range' and puts result into 'result'
		/// </summary>
		public void Sum(ExcelWorksheet sheet, ExcelCellAddress result, ExcelRange range) {
			sheet.Cells[result.Address].Formula = "SUM(" + range.Start.Address + ':' + range.End.Address + ")";
		}

		/// <summary>
		/// Sums cells in scattered range 'addresses' and puts result into 'result'
		/// </summary>
		public void Sum(ExcelWorksheet sheet, ExcelCellAddress result, ExcelCellAddress[] addresses) {
			string s = "";
			foreach (ExcelCellAddress addr in addresses) {
				s = string.Join(",", s, addr.Address);
			}
			s = s.TrimStart(',');
			sheet.Cells[result.Address].Formula = "SUM(" + s + ")";
		}

		/// <summary>
		/// Averages cells in 'range' and puts result into 'result'
		/// </summary>
		public void Average(ExcelWorksheet sheet, ExcelCellAddress result, ExcelRange range) {
			sheet.Cells[result.Address].Formula = "AVERAGE(" + range + ")";
		}

		/// <summary>
		/// Averages cells in scattered range 'addresses' and puts result into 'result'
		/// </summary>
		public void Average(ExcelWorksheet sheet, ExcelCellAddress result, ExcelCellAddress[] addresses) {
			string s = "";
			foreach (ExcelCellAddress addr in addresses) {
				s = string.Join(",", s, addr.Address);
			}
			s = s.TrimStart(',');
			sheet.Cells[result.Address].Formula = "AVERAGE(" + s + ")";
		}

		/// <summary>
		/// Preforms a Sum of 'numerator' and divides it by a value in 'denominator' puts the result into 'result'
		/// </summary>
		public void DivideBy(ExcelWorksheet sheet, ExcelCellAddress result, ExcelRange numerator, ExcelCellAddress denominator) {
			sheet.Cells[result.Address].Formula = "SUM(" + numerator.Start.Address + ':' + numerator.End.Address + ")/" + denominator.Address;
		}

		/// <summary>
		/// Gets continuous range between two addressess
		/// </summary>
		public ExcelRange GetRangeContinuous(ExcelWorksheet sheet, ExcelCellAddress addr1, ExcelCellAddress addr2) {
			sheet.Select(addr1.Address + ":" + addr2.Address);
			ExcelRange range = sheet.SelectedRange;
			sheet.Select("A1");
			return range;
		}

		/// <summary>
		/// Gets continuous range between two addressess as strings
		/// </summary>
		public ExcelRange GetRangeContinuous(ExcelWorksheet sheet, string addr1, string addr2) {
			sheet.Select(addr1 + ":" + addr2);
			ExcelRange range = sheet.SelectedRange;
			sheet.Select("A1");
			return range;
		}
		#endregion
	}
}
