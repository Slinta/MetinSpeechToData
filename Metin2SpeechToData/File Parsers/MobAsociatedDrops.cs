using System;
using System.Collections.Generic;
using System.IO;

namespace Metin2SpeechToData {
	public class MobAsociatedDrops {
		public const string MOB_DROPS_FILE = "Mob Asociated Drops.definition";
		ushort changeAcc = 0;

		/// <summary>
		/// Get all drops, entries are 'main pronounciation' of items in definition file
		/// </summary>
		public string[] getAllDropsFile { get; private set; }

		public MobAsociatedDrops() {

			try {
				getAllDropsFile = File.ReadAllLines(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MOB_DROPS_FILE);
			}
			catch {
				using(StreamWriter sw = File.CreateText(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MOB_DROPS_FILE)) {
					Console.WriteLine("Creating File for drop asociation to enemies...");
				}
				getAllDropsFile = File.ReadAllLines(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MOB_DROPS_FILE);
			}
		}

		~MobAsociatedDrops() {
			SaveChanges();
		}


		public void UpdateDrops(string mobName, string itemName) {
			bool mobEntryExists = false;
			for (int i = 0; i < getAllDropsFile.Length; i++) {
				if (getAllDropsFile[i].Contains("{")) {
					string[] splt = getAllDropsFile[i].Split('{');
					if (splt[0] == mobName) {
						mobEntryExists = true;
						bool edit = !CheckItemExists(getAllDropsFile[i + 1], itemName);
						if (edit) {
							getAllDropsFile[i + 1] = getAllDropsFile[i + 1] + "," + itemName;
							Program.interaction.AddItemEntryToCurrentSheet(itemName);
						}
						changeAcc++;
						if (changeAcc >= 10) {
							SaveChanges();
							changeAcc = 0;
						}
						return;
					}
				}
			}
			if (!mobEntryExists) {
				AddMobEntry(mobName, itemName);
			}
		}

		private void AddMobEntry(string mobName, string itemName) {
			List<string> modified = new List<string>(getAllDropsFile);
			modified.Add(mobName + "{");
			modified.Add("\t" + itemName);
			modified.Add("}");
			getAllDropsFile = modified.ToArray();
			changeAcc++;
		}


		public bool RemoveItemEntry(string mobName, string itemName) {
			bool yes = Configuration.GetBoolInput("Remove from current enemy's (" + Program.enemyHandling.getCurrentEnemy + ") item list?");
			if (yes) {
				List<string> modified = new List<string>(getAllDropsFile);
				for (int i = 0; i < modified.Count; i++) {
					if (modified[i].Contains(mobName)) {
						string[] split = modified[i + 1].Split(',');
						for (int j = 0; j < split.Length; j++) {
							if (split[j] == itemName) {
								split[j] = "";
								modified[i + 1] = string.Join(",", split);
								return true;
							}
						}
					}
				}
				return false;
			}
			else {
				Console.WriteLine("Entry not removed!");
				return false;
			}
		}

		private bool CheckItemExists(string line, string itemName) {
			string[] split = line.Split(',');
			for (int i = 0; i < split.Length; i++) {
				split[i] = split[i].Trim('\n', ' ', '\t');
				if (split[i] == itemName) {
					return true;
				}
			}
			return false;
		}

		public string[] GetDropsForMob(string enemyName) {
			for (int i = 0; i < getAllDropsFile.Length; i++) {
				if (getAllDropsFile[i].Contains("{")) {
					string mob = getAllDropsFile[i].Split('{')[0];
					if(mob == enemyName) {
						return getAllDropsFile[i + 1].Trim('\t', '\n', ' ').Split(',');
					}
				}
			}
			return null;
		}

		private void SaveChanges() {
			File.WriteAllLines(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MOB_DROPS_FILE, getAllDropsFile);
		}
	}
}
