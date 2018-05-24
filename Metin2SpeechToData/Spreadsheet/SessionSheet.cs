using OfficeOpenXml;
using System.IO;
using System;
using System.Collections.Generic;
using static Metin2SpeechToData.Spreadsheet.SsConstants;

namespace Metin2SpeechToData.Spreadsheet {
	class SessionSheet {

		public const string ENEMY_KILLS = "J5";
		public const string AVERAGE_KILL_REWARD = "J6";
		public const string MOST_COMMON_ENEMY = "J7";
		public const string AVERAGE_TIME_BETWEEN_KILLS = "J8";
		public const string SESSION_DURATION = "J9";
		public const string TOTAL_ITEM_VALUE = "P5";
		public const string GROUP_NO = "P6";
		public const string ITEM_NO = "P7";
		public const string AVG_EARN = "P8";
		public const string MOST_COMMON_ITEM = "P9";

		public ExcelWorksheet current { get; }

		private ExcelCellAddress currFreeAddress;

		private Data data;

		public SessionSheet() {
			string sessionsDir = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Sessions" + Path.DirectorySeparatorChar;
			string fileName = "_" + DateTime.Now.ToLongDateString() + ".xlsx";
			File.Create(sessionsDir + fileName);
			ExcelPackage package = new ExcelPackage(new FileInfo(sessionsDir + fileName));
			SpreadsheetTemplates template = new SpreadsheetTemplates();
			current = template.InitSessionSheet(package.Workbook);
			currFreeAddress = new ExcelCellAddress(ITEM_ROW, ITEM_COL);
			data = new Data();
		}

		public void Add(string itemName, string group, string enemy, DateTime dropTime, uint value) {
			current.SetValue(currFreeAddress.Address, itemName);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 4), enemy);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 7), group);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 10), dropTime.ToShortTimeString());
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 13), value);
			currFreeAddress = SpreadsheetHelper.OffsetAddress(currFreeAddress, 1, 0);
			data.UpdateData(itemName, group, enemy, dropTime, value);
		}

		public void Finish() {
			PopulateHeadder(data);
		}

		private void PopulateHeadder(Data data) {

		}

		private sealed class Data{
			public DateTime start { get; }
			public uint enemiesKilled { get; set; }
			public Dictionary<string, int> commonEnemy { get; }
			public Dictionary<string, int> commonItem { get; }
			public List<DateTime> dtopTimes { get; }
			public Dictionary<string, int> items { get; }
			public List<string> currentGroups { get; }

			public Data() {
				start = DateTime.Now;
			}

			public void UpdateData(string itemName, string group, string enemy, DateTime dropTime, uint value) {

			}
		}
	}
}
