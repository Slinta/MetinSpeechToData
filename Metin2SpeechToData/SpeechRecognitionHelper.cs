using System;
using System.Linq;
using System.Collections.Generic;
using System.Speech.Recognition;
using System.IO;

namespace Metin2SpeechToData {
	public class SpeechRecognitionHelper {

		public enum ModifierWords {
			NONE,
			NEW_TARGET,
			UNDO
		};

		/// <summary>
		/// Modifiers dictionary, used to convert enum valuies to spoken word
		/// </summary>
		public static IReadOnlyDictionary<ModifierWords, string> modifierDict = new Dictionary<ModifierWords, string>() {
			{ ModifierWords.NONE, "None" },
			{ ModifierWords.NEW_TARGET , "New Target" },
			{ ModifierWords.UNDO, "Undo" },
		};

		private SpeechRecognitionEngine control;
		private SpeechRecognitionEngine main;
		public ControlSpeechCommands controlCommands { get; private set; }

		public static ModifierWords currentModifier = ModifierWords.NONE;

		public event Program.Recognition OnRecognitionChange;

		public SpeechRecognitionHelper(ref SpeechRecognitionEngine engine) {
			main = engine;
			control = new SpeechRecognitionEngine();
			Grammar controlGrammar = LoadControlGrammar(out ControlSpeechCommands c);
			controlGrammar.Name = "Controler Grammar";
			controlCommands = c;

			control.LoadGrammar(controlGrammar);
			control.SetInputToDefaultAudioDevice();
			control.SpeechRecognized += Control_SpeechRecognized;
			if (Program.debug) {
				Console.WriteLine("Control grammar loaded...");
			}
			Console.Write("Available commands:\n" +
						  controlCommands.getStartCommand + " - Start recognition\n" +
						  controlCommands.getPauseCommand + " - Pauses main recognition\n" +
						  controlCommands.getStopCommand + " - Exits App\n" +
						  controlCommands.getSwitchGrammarCommand + " - Changes grammar (your drop location)\n");
			control.RecognizeAsync(RecognizeMode.Multiple);
		}


		private void Control_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			string res = e.Result.Text;

			if (res == controlCommands.getStartCommand) {
				Console.Write("Starting Recognition. Current grammar: ");
				if (main.Grammars.Count == 0) {
					Console.WriteLine("NOT INITIALIZED!");
					Console.WriteLine("Set grammar first with 'Computer switch grammar'");
					return;
				}
				else {
					Console.WriteLine(main.Grammars[0].Name);
					main.Grammars[0].Enabled = true;
					OnRecognitionChange(Program.RecognitionState.DROP_LOGGER_RUNNING);
				}
			}
			else if (res == controlCommands.getStopCommand) {
				Console.WriteLine("Stopping Recognition!");
				Environment.Exit(0);
			}
			else if (res == controlCommands.getPauseCommand) {
				Console.WriteLine("Pausing Recognition!");
				OnRecognitionChange(Program.RecognitionState.CONTROL_RUNNING);

			}
			else if (res == controlCommands.getSwitchGrammarCommand) {
				Choices definitions = new Choices();
				Console.Write("Switching Grammar, available: ");
				string[] available = DefinitionParser.instance.getDefinitionNames;
				for (int i = 0; i < available.Length; i++) {
					definitions.Add(available[i]);
					if (i == available.Length - 1) {
						Console.Write(available[i]);
					}
					else {
						Console.Write(available[i] + ", ");
					}
				}
				if (control.Grammars.Count == 1) {
					control.LoadGrammar(new Grammar(definitions));
				}
				control.Grammars[0].Enabled = false;
				control.SpeechRecognized -= Control_SpeechRecognized;
				control.SpeechRecognized += Switch_WordRecognized;
				OnRecognitionChange(Program.RecognitionState.SWITCHING);
			}
			else {
				OnRecognitionChange(Program.RecognitionState.ERROR);
			}
		}

		private Grammar LoadControlGrammar(out ControlSpeechCommands commands) {
			DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
			FileInfo grammarFile;
			try {
				grammarFile = dir.GetFiles("Control.definition")[0];
			}
			catch {
				throw new CustomException("Could not locate 'Control.definition' file! You have to redownload this application");
			}
			Choices choices = new Choices();
			using (StreamReader sr = grammarFile.OpenText()) {
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

		private void Switch_WordRecognized(object sender, SpeechRecognizedEventArgs e) {
			Console.WriteLine("\nSelected - " + e.Result.Text);
			main.UnloadAllGrammars();
			main.LoadGrammar(DefinitionParser.instance.GetGrammar(e.Result.Text));
			main.LoadGrammar(new Grammar(new Choices(modifierDict.Values.ToArray())));
			main.LoadGrammar(DefinitionParser.instance.GetMobGrammar("Mob_" + e.Result.Text));
			main.Grammars[0].Enabled = true;
			main.Grammars[1].Enabled = true;
			main.Grammars[2].Enabled = false;
			control.SpeechRecognized -= Switch_WordRecognized;
			control.Grammars[0].Enabled = true;
			control.SpeechRecognized += Control_SpeechRecognized;
			control.Grammars[1].Enabled = false;
			DefinitionParser.instance.currentGrammarFile = DefinitionParser.instance.GetDefinitionByName(e.Result.Text);
			DefinitionParser.instance.currentMobGrammarFile = DefinitionParser.instance.GetMobDefinitionByName("Mob_" + e.Result.Text);
			Program.interaction.MakeANewSpreadsheet(DefinitionParser.instance.currentGrammarFile);
			if (Program.debug) {
				Console.WriteLine(main.Grammars.Count);
			}
		}
	}
}
