using OfficeOpenXml;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using static Metin2SpeechToData.Spreadsheet.SsConstants;

namespace Metin2SpeechToData {
	public class SessionSheet {

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

		private readonly Data data;

		public SessionSheet(string name) {
			string sessionsDir = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Sessions" + Path.DirectorySeparatorChar;
			string fileName = "_" + DateTime.Now.ToLongDateString() + ".xlsx";
			File.Create(sessionsDir + fileName);
			ExcelPackage package = new ExcelPackage(new FileInfo(sessionsDir + fileName));
			SpreadsheetTemplates template = new SpreadsheetTemplates();
			current = template.InitSessionSheet(package.Workbook);
			currFreeAddress = new ExcelCellAddress(ITEM_ROW, ITEM_COL);
			current.SetValue(SsControl.C_SHEET_NAME, "Session:(_) " + DateTime.Now.ToLongDateString());
			data = new Data();
		}

		public void NewEnemy(string enemy, DateTime meetTime) {
			data.UpdateDataEnemy(enemy, false, meetTime);
		}

		public void EnemyKilled(string enemy, DateTime killTime) {
			data.UpdateDataEnemy(enemy, true, killTime);
		}

		public void Add(string itemName, string group, string enemy, DateTime dropTime, uint value) {
			current.SetValue(currFreeAddress.Address, itemName);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 4), enemy);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 7), group);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 10), dropTime.ToShortTimeString());
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 13), value);
			currFreeAddress = SpreadsheetHelper.OffsetAddress(currFreeAddress, 1, 0);
			data.UpdateDataItem(itemName, enemy, dropTime);
		}

		public void Finish() {
			PopulateHeadder(data);
		}

		private void PopulateHeadder(Data data) {
			current.SetValue(ENEMY_KILLS, data.enemiesKilled);
			current.SetValue(AVERAGE_KILL_REWARD, ""/*how even*/); //TODO average kill reward
			current.SetValue(MOST_COMMON_ENEMY, data.GetMostCommon(data.items));
			current.SetValue(AVERAGE_TIME_BETWEEN_KILLS, data.GetAvgTimeToKill());
			current.SetValue(TOTAL_ITEM_VALUE, data.items.Values.Sum());
			current.SetValue(GROUP_NO, data.currentGroups.Count);
			current.SetValue(ITEM_NO, data.items.Count);
			current.SetValue(AVG_EARN, data.GetAvgRewardFromEnemy()/*...*/); //TODO average earn per _TIME_PERIOD
			current.SetValue(MOST_COMMON_ITEM, data.GetMostCommon(data.commonEnemy));
		}

		private sealed class Data {
			public DateTime start { get; }
			public uint enemiesKilled { get; set; }
			public Dictionary<string, int> commonEnemy { get; }
			public List<DateTime> dropTimes { get; }
			public Dictionary<string, int> items { get; }
			public List<string> currentGroups { get; }

			public Data() {
				start = DateTime.Now;
				commonEnemy = new Dictionary<string, int>();
				dropTimes = new List<DateTime>();
				items = new Dictionary<string, int>();
				currentGroups = new List<string>();
			}

			public void UpdateDataItem(string itemName, string enemy, DateTime dropTime) {
				if (!items.ContainsKey(itemName)) {
					items.Add(itemName, 1);
				}
				else {
					items[itemName] += 1;
				}
				dropTimes.Add(dropTime);
				string group = DefinitionParser.instance.currentGrammarFile.GetGroup(itemName);
				if (!currentGroups.Contains(group)) {
					currentGroups.Add(group);
				}
			}


			public void UpdateDataEnemy(string enemyName, bool killed, DateTime dropTime) {
				if (killed) {
					enemiesKilled += 1;
					if (!commonEnemy.ContainsKey(enemyName)) {
						commonEnemy.Add(enemyName, 1);
					}
					else {
						commonEnemy[enemyName] += 1;
					}
				}
			}

			public string GetMostCommon(Dictionary<string, int> dict) {
				int max = 0;
				string name = "";
				foreach (KeyValuePair<string, int> item in dict) {
					if (item.Value < max) {
						max = item.Value;
						name = item.Key;
					}
				}
				return name + string.Format("({0})", max);
			}

			public string GetAvgTimeToKill() {
				return "";
			}

			public float GetAvgRewardFromEnemy() {
				TimeSpan step = Configuration.minutesAverageDropValueInterval;
				return 1f;
			}
		}
	}
}
