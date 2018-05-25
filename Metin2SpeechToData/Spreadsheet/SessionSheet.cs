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

		private readonly ExcelPackage package;

		public SessionSheet(string name) {
			string sessionsDir = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Sessions" + Path.DirectorySeparatorChar;
			string fileName = "_" + DateTime.Now.ToLongDateString() + "_" + DateTime.Now.ToLongTimeString().Replace(':', '-') + ".xlsx";
			package = new ExcelPackage(new FileInfo(sessionsDir + fileName));
			SpreadsheetTemplates template = new SpreadsheetTemplates();
			current = template.InitSessionSheet(package.Workbook);
			currFreeAddress = new ExcelCellAddress(ITEM_ROW, ITEM_COL);
			current.SetValue(SsControl.C_SHEET_NAME, "Session:(_) " + DateTime.Now.ToLongDateString());
			data = new Data();
			package.Save();
		}

		public void NewEnemy(string enemy, DateTime meetTime) {
			data.UpdateDataEnemy(enemy, false, meetTime);
		}

		public void EnemyKilled(string enemy, DateTime killTime) {
			data.UpdateDataEnemy(enemy, true, killTime);
		}

		public void Add(string itemName, string group, string enemy, DateTime dropTime, uint value) {
			current.SetValue(currFreeAddress.Address, itemName);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 7), enemy == "" ? UNSPEICIFIED_ENEMY : enemy);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 4), group);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 10), dropTime.ToShortTimeString());
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 13), value);
			currFreeAddress = SpreadsheetHelper.OffsetAddress(currFreeAddress, 1, 0);
			data.UpdateDataItem(itemName, enemy == "" ? UNSPEICIFIED_ENEMY : enemy, dropTime);
			package.Save();
		}
		public void Add(DefinitionParserData.Item item, string enemy, DateTime dropTime) {
			Add(item.mainPronounciation, item.group, enemy, dropTime, item.yangValue);
		}

		public void Finish() {
			PopulateHeadder(data);
		}

		private void PopulateHeadder(Data data) {
			current.SetValue(ENEMY_KILLS, data.enemiesKilled);
			current.SetValue(AVERAGE_KILL_REWARD, data.totalValue / data.enemiesKilled);
			current.SetValue(MOST_COMMON_ENEMY, data.GetMostCommonEntity(data.commonEnemy));
			current.SetValue(AVERAGE_TIME_BETWEEN_KILLS, TimeSpan.FromSeconds(data.GetAverageTimeBetweenInSeconds(data.enemyKillTimes)));
			current.SetValue(SESSION_DURATION, DateTime.Now.Subtract(data.start));
			current.SetValue(TOTAL_ITEM_VALUE, data.totalValue);
			current.SetValue(GROUP_NO, data.currentGroups.Count);
			current.SetValue(ITEM_NO, data.itemDropTimes.Count);
			current.SetValue(AVG_EARN, (data.totalValue / DateTime.Now.Subtract(data.start).TotalMinutes) * Configuration.minutesAverageDropValueInterval.TotalMinutes);
			current.SetValue(MOST_COMMON_ITEM, data.GetMostCommonEntity(data.items));
			package.Save();
		}

		private sealed class Data {
			public DateTime start { get; }
			public uint enemiesKilled { get; set; }
			public Dictionary<string, int> commonEnemy { get; }
			public List<DateTime> enemyKillTimes { get; }
			public List<DateTime> itemDropTimes { get; }
			public Dictionary<string, int> items { get; }
			public List<string> currentGroups { get; }
			public uint totalValue { get; set; }

			public Data() {
				start = DateTime.Now;
				commonEnemy = new Dictionary<string, int>();
				enemyKillTimes = new List<DateTime>();
				itemDropTimes = new List<DateTime>();
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
				itemDropTimes.Add(dropTime);
				string itemGroup = DefinitionParser.instance.getDefinitions[0].GetGroup(itemName);
				if (!currentGroups.Contains(itemGroup)) {
					currentGroups.Add(itemGroup);
				}
				totalValue += DefinitionParser.instance.getDefinitions[0].GetYangValue(itemName);

			}
			public void UpdateDataEnemy(string enemyName, bool killed, DateTime enemyActionTime) {
				if (killed) {
					enemiesKilled += 1;
					if (!commonEnemy.ContainsKey(enemyName)) {
						commonEnemy.Add(enemyName, 1);
					}
					else {
						commonEnemy[enemyName] += 1;
					}
					enemyKillTimes.Add(enemyActionTime);
				}
			}

			public string GetMostCommonEntity(Dictionary<string, int> dict) {
				int mostCommon = -1;
				string mostCommons = "";
				foreach (string key in dict.Keys) {
					if (dict[key] > mostCommon) {
						mostCommon = dict[key];
						mostCommons = key;
					}
					if (dict[key] == mostCommon) {
						mostCommons += (", " + key);
					}
				}
				return mostCommons;
			}

			public float GetAverageTimeBetweenInSeconds(List<DateTime> list) {
				double totalSeconds = 0;
				for (int i = 0; i < list.Count; i++) {
					if (i == 0) {
						totalSeconds += (list[0].Subtract(start)).TotalSeconds;
					}
					else {
						totalSeconds += (list[i].Subtract(list[i - 1])).TotalSeconds;
					}
				}
				return (float)(totalSeconds / list.Count);
			}
		}
	}
}
