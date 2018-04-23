using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Speech.Recognition;

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

		#region Constructor/Destructor
		public DefinitionParser() {
			DirectoryInfo d = new DirectoryInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Definitions");
			FileInfo[] filesPresent = d.GetFiles("*.definition",SearchOption.AllDirectories).Where(
				name => !name.Name.Split('.')[0].StartsWith("Control") &&
						!name.Name.Split('.')[0].StartsWith("Mob")).ToArray();
			if (filesPresent.Length == 0) {
				throw new CustomException("Your program is missing voice recognition strings! Either redownload, or create your own *.definition text file.");
			}
			if(d.GetFiles("Mob_*.definition").Length != 0) {
				getMobDefinitions = new MobParserData().Parse(d);
			}

			getDefinitions = new DefinitionParserData[filesPresent.Length];
			for (int i = 0; i < filesPresent.Length; i++) {
				DefinitionParserData data = new DefinitionParserData();
				using (StreamReader s = filesPresent[i].OpenText()) {
					data.ID = filesPresent[i].Name.Split('.')[0];
					data.groups = ParseHeader(s);
					data.entries = ParseEntries(s);
					data.ConstructGrammar();
					getDefinitions[i] = data;
				}
			}
			instance = this;
		}

		~DefinitionParser() {
			instance = null;
		}
		#endregion

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

		private DefinitionParserData.Entry[] ParseEntries(StreamReader r) {
			List<DefinitionParserData.Entry> entries = new List<DefinitionParserData.Entry>();
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
				entries.Add(new DefinitionParserData.Entry {
					mainPronounciation = same[0],
					group = split[2].TrimStart(' '),
					yangValue = uint.Parse(split[1].TrimStart(' ')),
					ambiguous = same.Where(a => a != same[0]).ToArray()
				});
			}
			return entries.ToArray();
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
				if (def.ID == identifier) {
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
				if (data.ID == name) {
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
