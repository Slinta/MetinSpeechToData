using System;
using System.Collections.Generic;
using System.Speech.Recognition;
using System.IO;

namespace Metin2SpeechToData {
	class SpeechRecognitionHelper {

		private SpeechRecognitionEngine control;
		private SpeechRecognitionEngine main;
		public ControlSpeechCommands controlCommands { get; private set; }

		public SpeechRecognitionHelper(ref SpeechRecognitionEngine engine) {
			main = engine;
			control = new SpeechRecognitionEngine();
			ControlSpeechCommands c;
			Grammar controlGrammar = LoadControlGrammar(out c);
			controlGrammar.Name = "Controler Grammar";
			controlCommands = c;

			control.LoadGrammar(controlGrammar);
			control.SetInputToDefaultAudioDevice();
			control.SpeechRecognized += Control_SpeechMatch;
			Console.WriteLine("Control grammar loaded...");
			control.RecognizeAsync(RecognizeMode.Multiple);
		}

		private void Control_SpeechMatch(object sender, SpeechRecognizedEventArgs e) {
			string res = e.Result.Text;

			if(res == controlCommands.getStartCommand) {
				Console.WriteLine("Starting Recognition. Current grammar: " + (main.Grammars.Count == 0 ? "NOT INITIALIZED":main.Grammars[0].Name));
			}
			else if(res == controlCommands.getStopCommand) {
				Console.WriteLine("Stopping Recognition!");
			}
			else if (res == controlCommands.getPauseCommand) {
				Console.WriteLine("Pausing Recognition!");
			}
			else if (res == controlCommands.getSwitchGrammarCommand) {
				Console.WriteLine("Switching Grammar, available: ");
				foreach (string s in DefinitionParser.instance.getDefinitions) {
					Console.Write(s + ", ");
				}

			}

			/* Switch cases must be known at compile time /./
			 * switch (e.Result.Text) {
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
					Console.WriteLine("Switching Grammar, current mode");
					break;
				}
			}
			*/
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
			Choices choices = new Choices();
			using(StreamReader sr = grammarFile.OpenText()) {
				while (!sr.EndOfStream) {
					string line = sr.ReadLine();
					if (line.StartsWith("#") || string.IsNullOrWhiteSpace(line)) {
						continue;
					}
					string modified = line.Split(':')[1].Remove(0, 1);
					choices.Add(new Choices(modified.Split(',')));
				}
			}
			commands = new ControlSpeechCommands(grammarFile.Name);
			return new Grammar(choices);
		}

		/// <summary>
		/// Checks said string for matches in alternatives 
		/// </summary>
		private bool IsAnyOf(string original, string[] alrenatives) {
			foreach (string alt in alrenatives) {
				if(original == alt) {
					return true;
				}
			}
			return false;
		}
	}
}
