using System;
using System.IO;
using System.Speech.Recognition;

namespace Metin2SpeechToData {
	public class DefinitionParser {

		private Definition[] definitions;

		public DefinitionParser() {
			DirectoryInfo d = new DirectoryInfo(Directory.GetCurrentDirectory());
			FileInfo[] filesPresent = d.GetFiles("*.definition");

			if (filesPresent.Length == 0) {
				throw new Exception("Your program is missing voice recognition strings! Either redownload, or create your own *.definition text file.");
			}
			definitions = new Definition[filesPresent.Length];
			for (int i = 0; i < filesPresent.Length; i++) {
				GrammarBuilder gBuilder = new GrammarBuilder();
				using (StreamReader s = filesPresent[i].OpenText()) {
					while (!s.EndOfStream) {
						gBuilder.Append(new Choices(s.ReadLine().Split(',')));
					}
					definitions[i] = new Definition() { name = filesPresent[i].Name, grammar = new Grammar(gBuilder) };
				}
			}
		}

		public Grammar GetGrammar(string identifier) {
			foreach (Definition def in definitions) {
				if(def.name == identifier) {
					return def.grammar;
				}
			}
			throw new Exception("Grammar with identifier " + identifier + " does not exist!");
		}
	}

	struct Definition {
		public string name;
		public Grammar grammar;
	}
}
