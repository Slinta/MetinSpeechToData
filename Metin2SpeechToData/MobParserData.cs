﻿using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Speech.Recognition;

namespace Metin2SpeechToData {

	public class MobParserData {

		public string ID;
		public Enemy[] enemies;
		public Grammar grammar;

		public enum MobClass {
			COMMON,
			HALF_BOSS,
			BOSS,
			METEOR,
			SPECIAL
		}

		/// <summary>
		/// Parses all mob files that exists in current folder
		/// </summary>
		public MobParserData[] Parse(DirectoryInfo dir) {
			List<MobParserData> dataList = new List<MobParserData>();
			List<Enemy> mobs = new List<Enemy>();
			FileInfo[] mobFiles = dir.GetFiles("Mob_*.definition");
			for (int i = 0; i < mobFiles.Length; i++) {
				using (StreamReader sr = mobFiles[i].OpenText()) {
					MobParserData data = new MobParserData() {
						ID = mobFiles[i].Name.Split('.')[0],
					};
					while (!sr.EndOfStream) {
						string line = sr.ReadLine();
						if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) {
							line = sr.ReadLine();
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
						Enemy parsed = new Enemy() {
							mobMainPronounciation = same[0],
							ambiguous = same.Where(a => a != same[0]).ToArray(),
							mobLevel = ushort.Parse(split[1]),
							mobClass = ParseClass(split[2]),
							asociatedDrops = null
						};
						mobs.Add(parsed);

					}
					data.enemies = mobs.ToArray();
					data.grammar = ConstructGrammar(data.enemies);
					dataList.Add(data);
				}
			}
			return dataList.ToArray();
		}

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
			throw new Exception("Invalid Mob type " + s);
		}

		/// <summary>
		/// Gets main mob pronounciation by comparing ambiguities
		/// </summary>
		public string GetMainPronounciation(string calledAmbiguity) {
			foreach (Enemy enemy in enemies) {
				if(calledAmbiguity == enemy.mobMainPronounciation) {
					return calledAmbiguity; // == enemy.mainMobPronounciation
				}
				foreach (string ambiguity in enemy.ambiguous) {
					if (ambiguity == calledAmbiguity) {
						return enemy.mobMainPronounciation;
					}
				}
			}
			throw new Exception("No entry found, data was parsed incorrectly");
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
			public string mobMainPronounciation;
			public string[] ambiguous;
			public ushort mobLevel;
			public string[] asociatedDrops;
			public MobClass mobClass;
		}
	}
}