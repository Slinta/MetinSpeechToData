using System;
using System.Collections.Generic;
using System.Speech.Recognition;
using System.IO;

namespace Metin2SpeechToData {
	class SpeechRecognitionHelper {

		private SpeechRecognitionEngine control;
		public ControlSpeechCommands controlCommands { get; private set; }

		public SpeechRecognitionHelper(ref SpeechRecognitionEngine engine) {
			control = new SpeechRecognitionEngine();
			ControlSpeechCommands c;
			Grammar controlGrammar = LoadControlGrammar(out c);
			controlCommands = c;

			control.LoadGrammar(controlGrammar);
			control.SetInputToDefaultAudioDevice();
			control.SpeechRecognized += Control_SpeechMatch;
			Console.WriteLine("Control grammar loaded...");
		}

		private void Control_SpeechMatch(object sender, SpeechRecognizedEventArgs e) {
			switch (e.Result.Text) {
				case ControlSpeechCommands.START: {
					Console.WriteLine("Starting Recognition, current mode");
					break;
				}
				case ControlSpeechCommands.STOP: {
					Console.WriteLine("Stopping Recognition, current mode");
					break;
				}
				case ControlSpeechCommands.PAUSE: {
					Console.WriteLine("Pausing Recognition, current mode");
					break;
				}
				case ControlSpeechCommands.SWITCH: {
					Console.WriteLine("Switching Recognition, current mode");
					break;
				}
			}
		}

		/// <summary>
		/// Loads grammar for the controling speech recognizer
		/// </summary>
		private Grammar LoadControlGrammar(out ControlSpeechCommands commands) {
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
			commands = new ControlSpeechCommands(grammarFile.Name);
			return new Grammar(gBuilder);
		}
	}
}
