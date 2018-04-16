using System.Speech.Recognition;

namespace Metin2SpeechToData {
	public class DefinitionParserData {
		/// <summary>
		/// Name of this deinition --> file name that was used to parse this data
		/// </summary>
		public string ID;

		/// <summary>
		/// Groups defined at the top of the file
		/// </summary>
		public string[] groups;

		/// <summary>
		/// All items that are described in the file
		/// </summary>
		public Entry[] entries;

		/// <summary>
		/// Grammar created from all the item names and ambiguities
		/// </summary>
		public Grammar grammar;

		public struct Entry {
			public string mainPronounciation;
			public string[] ambiguous;
			public uint yangValue;
			public string group;
		}

		/// <summary>
		/// Gets main item pronounciation by comparing ambiguities
		/// </summary>
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
