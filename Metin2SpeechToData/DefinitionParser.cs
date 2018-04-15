using System;
using System.IO;
using System.Linq;
using System.Speech.Recognition;

namespace Metin2SpeechToData {
	public class DefinitionParser {

		public static DefinitionParser instance { get; private set; }

		private Grammar[] definitions; //Holds all loaded definitions that were present in Application folder

		public DefinitionParser() {
			DirectoryInfo d = new DirectoryInfo(Directory.GetCurrentDirectory());
			FileInfo[] filesPresent = d.GetFiles("*.definition").Where(name => !name.Name.Split('.')[0].StartsWith("Control")).ToArray();
			
			if (filesPresent.Length == 1) {
				throw new Exception("Your program is missing voice recognition strings! Either redownload, or create your own *.definition text file.");
			}
			definitions = new Grammar[filesPresent.Length];
			for (int i = 0; i < filesPresent.Length; i++) {
				if (filesPresent[i].Name == "Control.definition") {
					
					continue;
				}
				Choices choices = new Choices();
				using (StreamReader s = filesPresent[i].OpenText()) {
					while (!s.EndOfStream) {
						string line = s.ReadLine();
						if (line.StartsWith("#") || string.IsNullOrWhiteSpace(line)) {
							continue;
						}
						choices.Add(new Choices(line.Split(',')));
					}
					definitions[i] = new Grammar(choices) { Name = filesPresent[i].Name.Split('.')[0] };
				}
			}
			instance = this;
		}

		~DefinitionParser() {
			instance = null;
		}

		public Grammar GetGrammar(string identifier) {
			foreach (Grammar def in definitions) {
				if (def.Name == identifier) {
					return def;
				}
			}
			throw new Exception("Grammar with identifier " + identifier + " does not exist!");
		}

		/// <summary>
		/// Get all parsed definitions by name
		/// </summary>
		public string[] getDefinitions {
			get {
				string[] strings = new string[definitions.Length];
				for (int i = 0; i < definitions.Length; i++) {
					strings[i] = definitions[i].Name;
				}
				return strings;
			}
		}
	}
}
