using System;
using System.Collections.Generic;
using System.Speech.Recognition;
using System.IO;

namespace Metin2SpeechToData {
	class SpeechRecognitionHelper {

		private SpeechRecognitionEngine engine;

		public SpeechRecognitionHelper(ref SpeechRecognitionEngine engine, string[] definitions) {
			Grammar controlGrammar = LoadControlGrammar();
			Console.WriteLine("Control grammar loaded...");
			this.engine = engine;
			GrammarBuilder builder = new GrammarBuilder(new Choices(definitions));
			Grammar g = new Grammar(builder);
			this.engine.LoadGrammar(g);
			Console.WriteLine("Voice commands for Metin2 loaded...");
		}

		/// <summary>
		/// Switches grammar of the Speech Recognizer, used when changing drop locations
		/// </summary>
		/// <param name="definitions">as defined in DefinitionParser</param>
		public bool SwitchGrammar(string[] definitions) {
			GrammarBuilder builder = new GrammarBuilder(new Choices(definitions));
			Grammar g = new Grammar(builder);
			engine.LoadGrammar(g);
			Console.WriteLine("Grammar changed");
			return true;
		}

		/// <summary>
		/// Loads grammar for the controling speech recognizer
		/// </summary>
		private Grammar LoadControlGrammar() {
			DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
			FileInfo grammarFile;
			try {
				grammarFile = dir.GetFiles("Control.definition")[0];
			}
			catch {
				throw new Exception("Could not locate 'Control.definition' file! You have to redownload this application");
			}
			GrammarBuilder gBuilder = new GrammarBuilder();
			using(StreamReader sr = grammarFile.OpenText()) {
				while (!sr.EndOfStream) {
					string[] line = sr.ReadLine().Split(',');
					gBuilder.Append(new Choices(line));
				}
			}
			return new Grammar(gBuilder);
		}
	}
}
