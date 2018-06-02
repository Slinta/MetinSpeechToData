using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Speech.Recognition;
using System.Text.RegularExpressions;

namespace Metin2SpeechToData {
	public class DefinitionParser {

		/// <summary>
		/// Instance of this class
		/// </summary>
		public static DefinitionParser instance { get; private set; }

		/// <summary>
		/// Get all parsed definitions
		/// </summary>
		public DefinitionParserData[] getDefinitions { get; }

		/// <summary>
		/// Get all parsed mob definitions
		/// </summary>
		public MobParserData[] getMobDefinitions { get; }

		/// <summary>
		/// Current major grammar file
		/// </summary>
		public DefinitionParserData currentGrammarFile { get; private set; }
		/// <summary>
		/// Current major mob grammar file
		/// </summary>
		public MobParserData currentMobGrammarFile { get; private set; }

		/// <summary>
		/// Hotkey parser to assign saved hotkeys for items
		/// </summary>
		public HotkeyPresetParser hotkeyParser { get; private set; }

		/// <summary>
		/// The regex string that was used to select definition files
		/// </summary>
		public string regexMatchString { get; private set; }

		/// <summary>
		/// Whethet user selected custom made definition
		/// </summary>
		public bool custonDefinitionsLoaded { get; internal set; }


		#region Constructor
		/// <summary>
		/// Parser for .definition files, constructor parses selected .definition files in Definitions folder
		/// </summary>
		public DefinitionParser() {
			regexMatchString = "";
			instance = this;
			DirectoryInfo d = new DirectoryInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Definitions");
			FileInfo[] filesPresent = d.GetFiles("*.definition", SearchOption.AllDirectories);
			if (filesPresent.Length == 0) {
				throw new CustomException("Your program is missing voice recognition strings! Either redownload, or create your own *.definition text file.");
			}
			getDefinitions = LoadDefinitionData(filesPresent, out List<int> enemyDefinitionIndexes);
			getMobDefinitions = LoadMobDefinitionData(filesPresent, enemyDefinitionIndexes);
		}
		#endregion

		#region FileParsing base

		private DefinitionParserData[] LoadDefinitionData(FileInfo[] filesPresent, out List<int> mobIndexes) {
			List<int> enemyDefinitions = new List<int>();
			List<DefinitionParserData> definitions = new List<DefinitionParserData>();
			for (int i = 0; i < filesPresent.Length; i++) {
				if (filesPresent[i].Name.StartsWith("Mob_")) {
					enemyDefinitions.Add(i);
					continue;
				}

				using (StreamReader s = filesPresent[i].OpenText()) {
					DefinitionParserData data = new DefinitionParserData(filesPresent[i].Name.Split('.')[0], ParseHeader(s), ParseEntries(s));
					data.ConstructGrammar();
					definitions.Add(data);
				}
			}
			mobIndexes = enemyDefinitions;
			return definitions.ToArray();
		}

		private MobParserData[] LoadMobDefinitionData(FileInfo[] filesPresent, List<int> indexes) {
			if (indexes.Count != 0) {
				MobParserData[] mobData = new MobParserData().Parse(filesPresent, indexes);
				for (int i = 0; i < mobData.Length; i++) {
					for (int j = 0; j < mobData.Length; j++) {
						if (mobData[i].ID == "Mob_" + getDefinitions[j].ID) {
							getDefinitions[j].hasEnemyCompanionGrammar = true;
						}
					}
				}
				return mobData;
			}
			return new MobParserData[0];
		}

		private string[] ParseHeader(StreamReader r) {
			List<string> strings = new List<string>();
			string line = r.ReadLine();
			while (line != "}") {
				if (line.StartsWith("#") || line.Contains("{")) {
					line = r.ReadLine();
					continue;
				}
				strings.Add(line.Trim('\t'));
				line = r.ReadLine();
			}
			return strings.ToArray();
		}

		private DefinitionParserData.Item[] ParseEntries(StreamReader r) {
			List<DefinitionParserData.Item> entries = new List<DefinitionParserData.Item>();
			while (!r.EndOfStream) {
				string line = r.ReadLine();
				if (string.IsNullOrWhiteSpace(line)) {
					continue;
				}

				string[] split = line.Split(',');
				string[] same = split[0].Split('/');
				for (int i = 0; i < same.Length; i++) {
					same[i] = same[i].TrimStart(' ');
				}
				entries.Add(new DefinitionParserData.Item(same[0],
														  same.Where(a => a != same[0]).ToArray(),
														  uint.Parse(split[1].TrimStart(' ')),
														  split[2].TrimStart(' ')));
			}
			return entries.ToArray();
		}
		#endregion

		/// <summary>
		/// If exist, loads hotkeys for selected area
		/// </summary>
		public void LoadHotkeys(string area) {
			if (hotkeyParser == null) {
				hotkeyParser = new HotkeyPresetParser(area);
			}
			else {
				hotkeyParser.Load(area);
			}
		}

		/// <summary>
		/// Get grammar by its name (file name)
		/// </summary>
		public Grammar GetGrammar(string identifier) {
			foreach (DefinitionParserData def in getDefinitions) {
				if (def.ID == identifier) {
					return def.grammar;
				}
			}
			throw new CustomException("Grammar with identifier " + identifier + " does not exist!");
		}

		/// <summary>
		/// Get mob grammar by its name (file name)
		/// </summary>
		public Grammar GetMobGrammar(string identifier) {
			foreach (MobParserData def in getMobDefinitions) {
				if (def.ID == "Mob_" + identifier) {
					return def.grammar;
				}
			}
			throw new CustomException("Grammar with identifier " + identifier + " does not exist!");
		}

		/// <summary>
		/// Get drop definition by its name (file name)
		/// </summary>
		public DefinitionParserData GetDefinitionByName(string name) {
			foreach (DefinitionParserData data in getDefinitions) {
				if(data.ID == name) {
					return data;
				}
			}
			throw new CustomException("Definiton not found");
		}

		/// <summary>
		/// Get mob definition by its name (file name)
		/// </summary>
		public MobParserData GetMobDefinitionByName(string name) {
			foreach (MobParserData data in getMobDefinitions) {
				if (data.ID == "Mob_" + name) {
					return data;
				}
			}
			throw new CustomException("Definiton not found");
		}

		/// <summary>
		/// Updates currentGrammarFile with one named 'grammar' from available list
		/// </summary>
		/// <param name="grammar"></param>
		public void UpdateCurrents(string grammar) {
			currentGrammarFile = GetDefinitionByName(grammar);
			if (currentGrammarFile.hasEnemyCompanionGrammar) {
				currentMobGrammarFile = GetMobDefinitionByName(grammar);
			}
		}

		/// <summary>
		/// Get names of files that were used to create definitions
		///// </summary>
		public string[] getDefinitionNames {
			get {
				string[] names = new string[getDefinitions.Length];
				for (int i = 0; i < getDefinitions.Length; i++) {
					names[i] = getDefinitions[i].ID;
				}
				return names;
			}
		}

		public FileInfo getFileInfoFromId(string id) {
			DirectoryInfo d = new DirectoryInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Definitions");
			foreach (DefinitionParserData existing in getDefinitions) {
				if (Equals(id, existing.ID)) {
					
					return new FileInfo(d + Path.DirectorySeparatorChar.ToString() + id + ".definition");

				}
			}
			foreach (MobParserData existing in getMobDefinitions) {
				if (Equals(id, existing.ID)) {

					return new FileInfo(d + Path.DirectorySeparatorChar.ToString() + id + ".definition");

				}
			}
			throw new CustomException("No definition parser data with that id exists");

		}

		/// <summary>
		/// Adds a string to the current item.definition file
		/// </summary>
		/// <param name="entry">String in correct format</param>
		public void AddItemEntry(string entry, string newGroup) {
			FileInfo file = getFileInfoFromId(currentGrammarFile.ID);
			bool endsWithNewLine = false;
			using (StreamReader s = file.OpenText()) {
				string all = s.ReadToEnd();
				if (all[all.Length - 1].Equals('\n')) {
					endsWithNewLine = true;
				}
			}
			using (StreamWriter s = file.AppendText()) {
				if (!endsWithNewLine) {
					s.WriteLine();
				}
				s.WriteLine(entry);
				s.Close();
			}
			if (!currentGrammarFile.groups.Contains(newGroup)) {
				List<string> all = new List<string>();
				using (StreamReader s = file.OpenText()) {
					while (!s.EndOfStream) {
						string line = s.ReadLine();
						
						if (line.Contains('}')) {
							all.Add('\t' + newGroup);
						}
						all.Add(line);
					}

				}
				File.WriteAllLines(file.FullName, all.ToArray());
			}
		}

		/// <summary>
		/// Adds a string to the current mob.definition file
		/// </summary>
		/// <param name="entry">String in correct format</param>
		public void AddMobEntry(string entry) {
			FileInfo file = getFileInfoFromId(currentMobGrammarFile.ID);
			bool endsWithNewLine = false;
			using (StreamReader s = file.OpenText()) {
				string all = s.ReadToEnd();
				if (all[all.Length - 1].Equals('\n')) {
					endsWithNewLine = true;
				}
			}
			using (StreamWriter s = file.AppendText()) {
				if (!endsWithNewLine) {
					s.WriteLine();
				}
				s.WriteLine(entry);
				s.Close();
			}
		}
	}
}
