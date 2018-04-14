using System;
using System.IO;
using System.Speech.Recognition;

namespace Metin2SpeechToData {
	public class DefinitionParser {

		public Grammar getCurrentGrammar { get; private set; }

		public DefinitionParser() {
			DirectoryInfo d = new DirectoryInfo(Directory.GetCurrentDirectory());
			FileInfo[] filesPresent = d.GetFiles("*.definition");

			if (filesPresent.Length == 0) {
				throw new Exception("Your program is missing voice recognition strings! Either redownload, or create your own *.definition text file.");
			}
			GrammarBuilder gBuilder = new GrammarBuilder();
			for (int i = 0; i < filesPresent.Length; i++) {
				using (StreamReader s = filesPresent[i].OpenText()) {
					while (!s.EndOfStream) {
						gBuilder.Append(new Choices(s.ReadLine().Split(',')));
					}
				}
			}
			getCurrentGrammar = new Grammar(gBuilder);
		}

		public Grammar GetGrammar(string identifier) {
			GrammarBuilder gBuilder = new GrammarBuilder();
			using (StreamReader sr = File.OpenText(Directory.GetCurrentDirectory() + identifier + ".definition")) {
				while (!sr.EndOfStream) {
					gBuilder.Append(new Choices(sr.ReadLine().Split(',')));
				}
			}
			getCurrentGrammar = new Grammar(gBuilder);
			return getCurrentGrammar;
		}
	}
}
