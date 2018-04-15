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

		public void ContructGrammar() {
			Choices main = new Choices();
			foreach (Entry e in entries) {
				main.Add(e.mainPronounciation);
				foreach (string s in e.ambiguous) {
					main.Add(s);
				}
			}
			grammar = new Grammar(main);
		}
	}
}
