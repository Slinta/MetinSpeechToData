using System;
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
		/// Current major grammar file files
		/// </summary>
		public DefinitionParserData currentGrammarFile;
		/// <summary>
		/// Current major grammar file files
		/// </summary>
		public MobParserData currentMobGrammarFile;

		/// <summary>
		/// Hotkey parser to assign saved hotkeys for items
		/// </summary>
		public HotkeyPresetParser hotkeyParser;

		#region Constructor/Destructor
		/// <summary>
		/// Parser for .definition files, constructor parses all .definition files in Definitions folder
		/// </summary>
		public DefinitionParser(Regex searchPattern) {
			instance = this;
			DirectoryInfo d = new DirectoryInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Definitions");
			FileInfo[] filesPresent = d.GetFiles("*.definition", SearchOption.AllDirectories).Where(path => searchPattern.IsMatch(path.Name)).ToArray();
			if (filesPresent.Length == 0) {
				throw new CustomException("Your program is missing voice recognition strings! Either redownload, or create your own *.definition text file.");
			}

			List<int> Mob_indexes = new List<int>();
			List<DefinitionParserData> definitions = new List<DefinitionParserData>();
			for (int i = 0; i < filesPresent.Length; i++) {
				if (filesPresent[i].Name.StartsWith("Mob_")) {
					Mob_indexes.Add(i);
					continue;
				}
				DefinitionParserData data = new DefinitionParserData();
				using (StreamReader s = filesPresent[i].OpenText()) {
					data.ID = filesPresent[i].Name.Split('.')[0];
					data.groups = ParseHeader(s);
					data.entries = ParseEntries(s);
					data.ConstructGrammar();
					definitions.Add(data);
				}
			}
			
			getDefinitions = definitions.ToArray();

			if(Mob_indexes.Count != 0) {
				getMobDefinitions = new MobParserData().Parse(d);
			}
		}

		~DefinitionParser() {
			Console.WriteLine("Destructor of definition parser");
		}
		#endregion

		#region FileParsing base
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
				entries.Add(new DefinitionParserData.Item {
					mainPronounciation = same[0],
					group = split[2].TrimStart(' '),
					yangValue = uint.Parse(split[1].TrimStart(' ')),
					ambiguous = same.Where(a => a != same[0]).ToArray()
				});
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
		/// Get names of files that were used to create definitions
		/// </summary>
		public string[] getDefinitionNames {
			get {
				string[] names = new string[getDefinitions.Length];
				for (int i = 0; i < getDefinitions.Length; i++) {
					names[i] = getDefinitions[i].ID;
				}
				return names;
			}
		}
	}
}
