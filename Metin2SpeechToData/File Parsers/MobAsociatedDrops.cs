using System;
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
							getAllDropsFile = list.ToArray();
							Program.interaction.AddItemEntryToCurrentSheet(item);
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
				List<string> AllDropsFileList = new List<string>(getAllDropsFile);
				for (int i = 0; i < getAllDropsFile.Length; i++) {
					if (getAllDropsFile[i].Contains(mobName)) {
						string targetGroup = DefinitionParser.instance.currentGrammarFile.GetGroup(itemName);
						int itemLineIndex = GetGroupLine(i + 1, targetGroup);
						string[] split = getAllDropsFile[itemLineIndex].Split(':')[1].Split(',');
						if (split.Length != 1) {
							split[split.Length - 1] = "";
							getAllDropsFile[itemLineIndex] = "\t-" + targetGroup + ":" + string.Join(",", split.Where(str => str != ""));
							Console.WriteLine("Entry removed!");
						}
						else {
							AllDropsFileList.RemoveAt(itemLineIndex);
							getAllDropsFile = AllDropsFileList.ToArray();
						}
						SaveChanges();
						Program.interaction.RemoveItemEntryFromCurrentSheet(item);
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

		private bool CheckItemExists(string line, string itemName) {
			line = line.Split(':')[1];
			string[] split = line.Split(',');
			for (int i = 0; i < split.Length; i++) {
				if (split[i] == itemName) {
					return true;
				}
			}
			return false;
		}


		private bool GroupEntryExists(int mobLine, string group) {
			foreach (int line in GetLinesOfEnemy(mobLine)) {
				string s = getAllDropsFile[line].TrimStart().Trim('-').Split(':')[0];
				if (s == group) {
					return true;
				}
			}
			return false;
		}


		/// <summary>
		/// Returns all items this enemy drops, indexed by group name, updates dynamically.
		/// </summary>
		public Dictionary<string, string[]> GetDropsForMob(string enemyName) {
			Dictionary<string, string[]> ret = new Dictionary<string, string[]>();
			int[] lineIndexes = GetLinesOfEnemy(enemyName);
			foreach (int index in lineIndexes) {
				string currLine = getAllDropsFile[index];
				if (currLine.StartsWith("-")) {
					string[] split = currLine.Split(':');
					ret.Add(split[0].Trim('\t', '-'), split[1].Split(','));
				}
			}
			return ret;
		}

		private int[] GetLinesOfEnemy(int mobLineFirst) {
			List<int> lines = new List<int>() { mobLineFirst };
			string current = getAllDropsFile[mobLineFirst];
			while (!current.Contains('}')) {
				mobLineFirst++;
				lines.Add(mobLineFirst);
				current = getAllDropsFile[mobLineFirst];
			}
			return lines.ToArray();
		}

		private int[] GetLinesOfEnemy(string enemyName) {
			List<int> lines = new List<int>();
			if(getAllDropsFile.Length == 0) {
				return lines.ToArray();
			}

			int _index = 0;
			string currLine = getAllDropsFile[_index];
			while (!currLine.Contains(enemyName)) {
				_index++;
				if (_index == getAllDropsFile.Length) {
					return lines.ToArray();
				}
				currLine = getAllDropsFile[_index];
				
			}
			//Fond the start of our enemy
			lines.Add(_index);

			while (!currLine.Contains('}')) {
				lines.Add(++_index);
				currLine = getAllDropsFile[_index];
			}
			return lines.ToArray();
		}

		private int GetGroupLine(int enemyLine, string group) {
			foreach (int line in GetLinesOfEnemy(enemyLine)) {
				if (getAllDropsFile[line].Split(':')[0].TrimStart().TrimStart('-') == group) {
					return line;
				}
			}
			throw new CustomException("parameters not present in file");
		}

		public int GetGroupNumberForEnemy(string enemy, string group) {
			int[] linesOfEnemy = GetLinesOfEnemy(enemy);
			int i = linesOfEnemy.Length - 2;
			foreach (int line in linesOfEnemy) {
				if (getAllDropsFile[line].Contains(group)) {
					return i;
				}
				i--;
			}
			throw new CustomException("Such enemy has no such group");
		}

		private void SaveChanges() {
			File.WriteAllLines(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MOB_DROPS_FILE, getAllDropsFile);
		}
	}
}
