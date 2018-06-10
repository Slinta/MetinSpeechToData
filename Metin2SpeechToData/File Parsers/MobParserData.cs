using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Speech.Recognition;
using System;

namespace Metin2SpeechToData {

	public class MobParserData {

		public string ID { get; private set; }
		public Enemy[] enemies { get; private set; }
		public Grammar grammar { get; private set; }

		public enum MobClass {
			COMMON,
			HALF_BOSS,
			BOSS,
			METEOR,
			SPECIAL
		}

		/// <summary>
		/// Parses all mob files that exist in current folder
		/// </summary>
		public MobParserData[] Parse(FileInfo[] files, List<int> indexes) {
			List<MobParserData> dataList = new List<MobParserData>();
			List<Enemy> mobs = new List<Enemy>();
			for (int i = 0; i < indexes.Count; i++) {
				using (StreamReader sr = files[indexes[i]].OpenText()) {
					MobParserData data = new MobParserData() {
						ID = files[indexes[i]].Name.Split('.')[0],
					};
					while (!sr.EndOfStream) {
						string line = sr.ReadLine();
						if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) {
							continue;
						}
						string[] split = line.Split(',');
						string[] same = split[0].Split('/');
						for (int j = 0; j < same.Length; j++) {
							same[j] = same[j].Trim(' ');
						}
						for (int k = 0; k < split.Length; k++) {
							split[k] = split[k].Trim(' ');
						}
						Enemy parsed = new Enemy(same[0],
												 same.Where(a => a != same[0]).ToArray(),
												 ushort.Parse(split[1]),
												 ParseClass(split[2]));
						mobs.Add(parsed);
					}
					data.enemies = mobs.ToArray();
					data.grammar = ConstructGrammar(data);
					dataList.Add(data);
				}
			}
			return dataList.ToArray();
		}
		/// <summary>
		/// Parses all mob files that exist in current folder using only Mob_ prefiex files
		/// </summary>
		public MobParserData[] Parse(FileInfo[] files) {
			List<int> indexes = new List<int>();
			for (int i = 0; i < files.Length; i++) {
				indexes.Add(i);
			}
			return Parse(files, indexes);
		}

		/// <summary>
		/// Gets main mob pronounciation by comparing ambiguities
		/// </summary>
		public string GetMainPronounciation(string calledAmbiguity) {
			foreach (Enemy enemy in enemies) {
				if (calledAmbiguity == enemy.mobMainPronounciation) {
					return calledAmbiguity; // == enemy.mainMobPronounciation
				}
				foreach (string ambiguity in enemy.ambiguous) {
					if (ambiguity == calledAmbiguity) {
						return enemy.mobMainPronounciation;
					}
				}
			}
			throw new CustomException("No entry found, data was parsed incorrectly");
		}

		public MobClass ParseClass(string s) {
			s = s.Trim(' ');
			switch (s) {
				case "COMMON": {
					return MobClass.COMMON;
				}
				case "HALF_BOSS": {
					return MobClass.HALF_BOSS;
				}
				case "BOSS": {
					return MobClass.BOSS;
				}
				case "METEOR": {
					return MobClass.METEOR;
				}
				case "SPECIAL": {
					return MobClass.SPECIAL;
				}
			}
			throw new CustomException("Invalid Mob type " + s);
		}

		public Grammar ConstructGrammar(MobParserData data) {
			Choices main = new Choices();
			foreach (Enemy e in data.enemies) {
				main.Add(e.mobMainPronounciation);
				foreach (string s in e.ambiguous) {
					main.Add(s);
				}
			}
			return new Grammar(main) { Name = data.ID };
		}

		public void AddMobDuringRuntime(Enemy mob) {
			List<Enemy> enemyList = new List<Enemy>(enemies);
			foreach (Enemy entry in enemies) {
				if (entry.mobMainPronounciation == mob.mobMainPronounciation) {
					throw new CustomException("You can't add an item that already exists");
				}
			}
			enemyList.Add(mob);
			enemies = enemyList.ToArray();

			grammar = ConstructGrammar(this);
		}

		public struct Enemy {
			public Enemy(string mobMainPronounciation, string[] ambiguous, ushort mobLevel, MobClass mobClass) {
				this.mobMainPronounciation = mobMainPronounciation;
				this.ambiguous = ambiguous;
				this.mobLevel = mobLevel;
				this.mobClass = mobClass;
			}

			public string mobMainPronounciation { get; }
			public string[] ambiguous { get; }
			public ushort mobLevel { get; }
			public MobClass mobClass { get; }
		}
	}
}
