using System.Speech.Recognition;

namespace Metin2SpeechToData {
	public class DefinitionParserData {
		public string ID;
		public string[] groups;
		public Entry[] entries;
		public Grammar grammar;

		public struct Entry {
			public string mainPronounciation;
			public string[] ambiguous;
			public uint yangValue;
			public string group;
		}

		public string GetMainPronounciation(string calledAmbiguity) {
			foreach (Entry entry in entries) {
				foreach (string ambiguity in entry.ambiguous) {
					if (ambiguity == calledAmbiguity) {
						return entry.mainPronounciation;
					}
				}
			}

			return null;
		}

		public void ConstructGrammar() {
			Choices main = new Choices();
			foreach (Entry e in entries) {
				main.Add(e.mainPronounciation);
				foreach (string s in e.ambiguous) {
					main.Add(s);
				}
			}
			grammar = new Grammar(main) {
				Name = ID
			};
		}
	}
}
