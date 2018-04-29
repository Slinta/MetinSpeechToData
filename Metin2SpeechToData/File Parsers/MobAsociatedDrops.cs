﻿using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;

namespace Metin2SpeechToData {
	public class MobAsociatedDrops {

		public const string MOB_DROPS_FILE = "Mob Asociated Drops.definition";

		/// <summary>
		/// Get all drops, entries are 'main pronounciation' of items in definition file
		/// </summary>
		public string[] getAllDropsFile { get; private set; }

		public Dictionary<string, Dictionary<string, int>> numberForGroupForSheet = new Dictionary<string, Dictionary<string, int>>();

		public MobAsociatedDrops() {
			try {
				getAllDropsFile = File.ReadAllLines(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MOB_DROPS_FILE);
			}
			catch {
				using (StreamWriter sw = File.CreateText(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MOB_DROPS_FILE)) {
					Console.WriteLine("Creating File for drop asociation to enemies...");
				}
				getAllDropsFile = File.ReadAllLines(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MOB_DROPS_FILE);
			}
		}

		public void UpdateDrops(string mobName, DefinitionParserData.Item item) {
			if (!numberForGroupForSheet.ContainsKey(mobName)) {
				//First spotting of the enemy
				numberForGroupForSheet.Add(mobName, new Dictionary<string, int>() { { item.group, 0 } });
			}
			else if (!numberForGroupForSheet[mobName].ContainsKey(item.group)) {
				//First time seeing that group
				int newValue = numberForGroupForSheet[mobName].Count;
				numberForGroupForSheet[mobName].Add(item.group, newValue);
			}
			bool mobEntryExists = false;
			List<string> list;
			for (int i = 0; i < getAllDropsFile.Length; i++) {
				if (getAllDropsFile[i].Contains("{")) {
					string[] splt = getAllDropsFile[i].Split('{');
					if (splt[0] == mobName) {
						mobEntryExists = true;
						if (!GroupEntryExists(i + 1, item.group)) {
							list = new List<string>(getAllDropsFile);
							list.Insert(i + 1, "\t-" + item.group + ":");

							list[i + 1] = list[i + 1] + item.mainPronounciation;
							Program.interaction.AddItemEntryToCurrentSheet(item);
							getAllDropsFile = list.ToArray();
							SaveChanges();
							return;
						}

						int line = GetGroupLine(i + 1, item.group);
						bool edit = !CheckItemExists(getAllDropsFile[line], item.mainPronounciation);
						if (edit) {
							getAllDropsFile[line] = getAllDropsFile[line] + "," + item.mainPronounciation;
							Program.interaction.AddItemEntryToCurrentSheet(item);
							SaveChanges();
						}

						return;
					}
				}
			}
			if (!mobEntryExists) {
				AddMobEntry(mobName, item);
				Program.interaction.AddItemEntryToCurrentSheet(item);
				SaveChanges();
			}
		}

		private void AddMobEntry(string mobName, DefinitionParserData.Item item) {
			List<string> modified = new List<string>(getAllDropsFile);
			modified.Add(mobName + "{");
			modified.Add("\t-" + item.group + ":" + item.mainPronounciation);
			modified.Add("}");
			getAllDropsFile = modified.ToArray();
		}

		/// <summary>
		/// Prompts user to remove item from mob asociations, used when Undoing
		/// </summary>
		public bool RemoveItemEntry(string mobName, string itemName, bool yes) {
			DefinitionParserData.Item item = DefinitionParser.instance.currentGrammarFile.GetItemEntry(itemName);
			if (yes) {
				for (int i = 0; i < getAllDropsFile.Length; i++) {
					if (getAllDropsFile[i].Contains(mobName)) {
						string targetGroup = DefinitionParser.instance.currentGrammarFile.GetGroup(itemName);
						int itemLineIndex = GetGroupLine(i + 1, targetGroup);
						string[] split = getAllDropsFile[itemLineIndex].Split(':')[1].Split(',');
						split[split.Length - 1] = "";
						getAllDropsFile[itemLineIndex] = "\t-" + targetGroup + ":" + string.Join(",", split.Where(str => str != ""));
						Console.WriteLine("Entry removed!");
						SaveChanges();
						RemoveExcelSheetEntry(itemName);
						return true;
					}
				}
				Console.WriteLine("Entry not found?!");
				return false;
			}
			else {
				Console.WriteLine("Entry not removed!");
				return false;
			}
		}

		private void RemoveExcelSheetEntry(string itemName) {
			Program.interaction.RemoveItemEntryFromCurrentSheet(itemName);
		}

		private bool CheckItemExists(string line, string itemName) {
			line = line.Split(':')[1];
			string[] split = line.Split(',');
			for (int i = 0; i < split.Length; i++) {
				//TODO this could be removed ?
				split[i] = split[i].Trim('\n', ' ', '\t');
				if (split[i] == itemName) {
					return true;
				}
			}
			return false;
		}

		private bool GroupEntryExists(int mobLine, string group) {
			foreach (int line in GetLinesOfOneMob(mobLine)) {
				string s = getAllDropsFile[line].TrimStart().Trim('-').Split(':')[0];
				if (s == group) {
					return true;
				}
			}
			return false;
		}
		/// <summary>
		/// Returns all items this enemy drops, updates dynamically.
		/// </summary>
		public string[] GetDropsForMob(string enemyName) {
			for (int i = 0; i < getAllDropsFile.Length; i++) {
				if (getAllDropsFile[i].Contains("{")) {
					string mob = getAllDropsFile[i].Split('{')[0];
					if (mob == enemyName) {
						return getAllDropsFile[i + 1].Trim('\t', '\n', ' ').Split(',');
					}
				}
			}
			return new string[0];
		}

		private int[] GetLinesOfOneMob(int mobLineFirst) {
			List<int> lines = new List<int>() { mobLineFirst };
			string current = getAllDropsFile[mobLineFirst];
			while (!current.Contains('}')) {
				mobLineFirst++;
				lines.Add(mobLineFirst);
				current = getAllDropsFile[mobLineFirst];
			}
			return lines.ToArray();
		}

		private int GetGroupLine(int enemyLine, string group) {
			foreach (int line in GetLinesOfOneMob(enemyLine)) {
				if (getAllDropsFile[line].Split(':')[0].TrimStart().TrimStart('-') == group) {
					return line;
				}
			}


			throw new CustomException("parameters not present in file");
		}

		public int GetGroupNumberForEnemy(string enemy, string group) {
			return numberForGroupForSheet[enemy][group];
		}

		private void SaveChanges() {
			File.WriteAllLines(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MOB_DROPS_FILE, getAllDropsFile);
		}
	}
}
