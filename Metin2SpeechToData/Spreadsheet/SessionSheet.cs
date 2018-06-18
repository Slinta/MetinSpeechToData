using OfficeOpenXml;
using System.IO;
using System;
using System.Collections.Generic;
using static Metin2SpeechToData.Spreadsheet.SsConstants;

namespace Metin2SpeechToData {
	public class SessionSheet : IDisposable {

		public const string LINK_TO_MAIN = "B5";
		public const string SESSION_AREA_NAME = "B7";
		public const string MERGED_STATUS = "B9";
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
		private FileInfo mainFile { get; }
		private SpreadsheetInteraction interaction { get; }

		private ExcelCellAddress currFreeAddress;

		private readonly Data data;
		public ExcelPackage package { get; }

		

		public SessionSheet(SpreadsheetInteraction interaction, string name, FileInfo mainSheet) {
			DateTime dt = DateTime.Now;
			string fileName = $"Session {dt:dd.MM.yyyy}__{dt:hh}h-{dt:mm}m-{dt:ss}s.xlsx";

			package = new ExcelPackage(new FileInfo(Configuration.sessionDirectory + fileName));
			mainFile = mainSheet;
			SpreadsheetTemplates template = new SpreadsheetTemplates();
			current = template.InitSessionSheet(package.Workbook);
			currFreeAddress = new ExcelCellAddress(DATA_FIRST_ENTRY);
			data = new Data();
			this.interaction = interaction;
			current.SetValue(SESSION_AREA_NAME, name);
			current.SetValue(MERGED_STATUS, "Not merged!");
			package.Save();
		}

		private void PrepareRows(int rowCount) {
			current.Select(currFreeAddress.Address + ":" + SpreadsheetHelper.OffsetAddress(currFreeAddress, 0, 15).Address);
			ExcelRange yellow = current.SelectedRange;
			for (int i = 1; i <= rowCount; i++) {
				current.Select(SpreadsheetHelper.OffsetAddress(currFreeAddress, i, 0).Address);
				yellow.Copy(current.SelectedRange);
			}
		}

		/// <summary>
		/// Copies formatting form cells above and writes an entry to them, updates statistics data holders
		/// </summary>
		public void WriteOut() {
			PrepareRows(1);

			ItemMeta item = Undo.instance.itemInsertionList.Last.Value;
			Undo.instance.itemInsertionList.RemoveLast();

			current.SetValue(currFreeAddress.Address, item.itemBase.mainPronounciation);

			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 7), item.comesFromEnemy == "" ? UNSPEICIFIED_ENEMY : item.comesFromEnemy);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 4), item.itemBase.group);
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 10), item.dropTime.ToLongTimeString());
			current.SetValue(SpreadsheetHelper.OffsetAddressString(currFreeAddress, 0, 13), item.itemBase.yangValue);

			currFreeAddress = SpreadsheetHelper.OffsetAddress(currFreeAddress, 1, 0);

			data.UpdateDataItem(item);
			package.Save();
		}

		/// <summary>
		/// Finalizes session, writes remaining items in undo collections, fills headder
		/// </summary>
		public void Finish() {
			while (Undo.instance.itemInsertionList.Count != 0) {
				WriteOut();
			} 
			while(Undo.instance.enemyList.Count != 0) {
				Undo.Target enemy = Undo.instance.enemyList.Last.Value;
				Undo.instance.enemyList.RemoveLast();
				if (enemy.state == Undo.TargetStates.Killed) {
					data.UpdateDataEnemy(enemy.name, true, enemy.killTime);
				}
			}
			PopulateHeadder(data);


			Console.Write("This sessions name: ");
			string sessionName = Console.ReadLine();
			current.SetValue(SsControl.C_SHEET_NAME, "Session: " + sessionName + " | " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString());
			current.SetValue(SpreadsheetHelper.OffsetAddress(AVG_EARN, 0, -3).Address,
				(Configuration.minutesAverageDropValueInterval.TotalMinutes < 60 ?
				string.Format("Average Yang per {0} minutes", Configuration.minutesAverageDropValueInterval.TotalMinutes) : "Average Yang per hour"));
			SpreadsheetHelper.HyperlinkAcrossFiles(mainFile, H_DEFAULT_SHEET_NAME, interaction.UnmergedLinkSpot(sessionName), current, LINK_TO_MAIN, ">>Main Sheet<<");

			package.Save();
			package.File.Attributes = FileAttributes.Archive;
			package.Dispose();
		}

		private void PopulateHeadder(Data data) {
			current.SetValue(ENEMY_KILLS, data.enemiesKilled);
			if(data.enemiesKilled == 0) {
				current.SetValue(AVERAGE_KILL_REWARD, "Unavailable");
			}
			else {
				current.SetValue(AVERAGE_KILL_REWARD, data.totalValueFromEnemies / data.enemiesKilled);
			}
			current.SetValue(MOST_COMMON_ENEMY, data.GetMostCommonEntity(data.commonEnemy));
			current.SetValue(AVERAGE_TIME_BETWEEN_KILLS, TimeSpan.FromSeconds(data.GetAverageTimeBetweenInSeconds(data.enemyKillTimes)).ToString());
			current.SetValue(SESSION_DURATION, DateTime.Now.Subtract(data.start).ToString());
			current.SetValue(TOTAL_ITEM_VALUE, data.totalValue);
			current.SetValue(GROUP_NO, data.currentGroups.Count);
			current.SetValue(ITEM_NO, data.itemDropTimes.Count);
			current.SetValue(AVG_EARN, ((data.totalValue / DateTime.Now.Subtract(data.start).TotalMinutes) * Configuration.minutesAverageDropValueInterval.TotalMinutes));
			current.SetValue(MOST_COMMON_ITEM, data.GetMostCommonEntity(data.items));
		}


		private sealed class Data {

			public DateTime start { get; }
			public int enemiesKilled { get; set; }
			public Dictionary<string, int> commonEnemy { get; }
			public List<DateTime> enemyKillTimes { get; }
			public List<DateTime> itemDropTimes { get; }
			public Dictionary<string, int> items { get; }
			public List<string> currentGroups { get; }
			public uint totalValue { get; set; }
			public uint totalValueFromEnemies { get; set; }


			public Data() {
				start = DateTime.Now;
				commonEnemy = new Dictionary<string, int>();
				enemyKillTimes = new List<DateTime>();
				itemDropTimes = new List<DateTime>();
				items = new Dictionary<string, int>();
				currentGroups = new List<string>();
			}

			/// <summary>
			/// Parses item into internal statistics
			/// </summary>
			public void UpdateDataItem(ItemMeta item) {
				if (!items.ContainsKey(item.itemBase.mainPronounciation)) {
					items.Add(item.itemBase.mainPronounciation, 1);
				}
				else {
					items[item.itemBase.mainPronounciation] += 1;
				}
				itemDropTimes.Add(item.dropTime);
				string itemGroup = DefinitionParser.instance.getDefinitions[0].GetGroup(item.itemBase.mainPronounciation);
				if (!currentGroups.Contains(itemGroup)) {
					currentGroups.Add(itemGroup);
				}
				totalValue += DefinitionParser.instance.getDefinitions[0].GetYangValue(item.itemBase.mainPronounciation);
				if (item.comesFromEnemy != "") {
					totalValueFromEnemies += DefinitionParser.instance.getDefinitions[0].GetYangValue(item.itemBase.mainPronounciation);
				}
			}

			/// <summary>
			/// Parses 'enemyName' into internal statistics at given 'enemyActionTime' with killed/alive
			/// </summary>
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

			/// <summary>
			/// Returns key with highest value count, if multiple keys have the same value, they are appended
			/// </summary>
			public string GetMostCommonEntity(Dictionary<string, int> dict) {
				int mostCommon = -1;
				string name = "";
				foreach (string key in dict.Keys) {
					if (dict[key] > mostCommon) {
						mostCommon = dict[key];
						name = key;
						continue;
					}
					if (dict[key] == mostCommon) {
						name += (", " + key);
					}
				}
				return (name == "") ? "No encounters or drops ;)" : name;
			}


			public float GetAverageTimeBetweenInSeconds(List<DateTime> list) {
				if (list.Count == 0) {
					return 0;
				}
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

		public struct ItemMeta {
			public ItemMeta(DefinitionParserData.Item itemBase, string comesFromEnemy, DateTime dropTime, int amount) {
				this.itemBase = itemBase;
				this.comesFromEnemy = comesFromEnemy;
				this.dropTime = dropTime;
				this.amount = amount;
			}

			public DefinitionParserData.Item itemBase { get; }
			public string comesFromEnemy { get; }
			public DateTime dropTime { get; }
			public int amount { get; }
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				package.Dispose();
				disposedValue = true;
			}
		}

		 ~SessionSheet() {
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
