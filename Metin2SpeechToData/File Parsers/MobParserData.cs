using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Speech.Recognition;

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
												 GetAsociatedDrops(same[0]),
												 ParseClass(split[2]));
						mobs.Add(parsed);
					}
					data.enemies = mobs.ToArray();
					data.grammar = ConstructGrammar(data.enemies);
					dataList.Add(data);
				}
			}
			return dataList.ToArray();
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

		#region Mod drop file parser
		private string[] GetAsociatedDrops(string mobMainPronounciation) {
			string path = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Mob Asociated Drops.definition";
			if (File.Exists(path)) {
				using (StreamReader r = File.OpenText(path)) {
					while (!r.EndOfStream) {
						string line = r.ReadLine();
						if (line.StartsWith("#")) {
							continue;
						}
						if (line.Contains("{")) {
							string[] split = line.Split('{');
							if (split[0] == mobMainPronounciation) {
								return ParseMobDropLine(r);
							}
						}
					}
				}
			}
			return new string[0];
		}

		private string[] ParseMobDropLine(StreamReader r) {
			string[] drops = r.ReadLine().Split(',');
			for (int i = 0; i < drops.Length; i++) {
				drops[i] = drops[i].Trim('\n', ' ', '\t');
			}
			return drops;
		}
		#endregion

		private MobClass ParseClass(string s) {
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

		public Grammar ConstructGrammar(Enemy[] enemies) {
			Choices main = new Choices();
			foreach (Enemy e in enemies) {
				main.Add(e.mobMainPronounciation);
				foreach (string s in e.ambiguous) {
					main.Add(s);
				}
			}
			return new Grammar(main) { Name = ID };
		}

		public struct Enemy {
			public Enemy(string mobMainPronounciation, string[] ambiguous, ushort mobLevel, string[] asociatedDrops, MobClass mobClass) {
				this.mobMainPronounciation = mobMainPronounciation;
				this.ambiguous = ambiguous;
				this.mobLevel = mobLevel;
				this.asociatedDrops = asociatedDrops;
				this.mobClass = mobClass;
			}

			public string mobMainPronounciation { get; }
			public string[] ambiguous { get; }
			public ushort mobLevel { get; }
			public string[] asociatedDrops { get; }
			public MobClass mobClass { get; }
		}
	}
}
