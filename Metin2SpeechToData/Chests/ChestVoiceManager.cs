using System;
using System.Collections.Generic;
using System.Speech.Recognition;

namespace Metin2SpeechToData.Chests {
	class ChestVoiceManager : SpeechRecognitionHelper {

		private SpeechRecognitionEngine main;
		private SpeechRecognitionEngine game;
		private DefinitionParser parser;
		
		#region Contrsuctor / Destructor
		public ChestVoiceManager(ref SpeechRecognitionEngine engine) {
			game = engine;
			main = new SpeechRecognitionEngine();
			parser = new DefinitionParser(new System.Text.RegularExpressions.Regex(@"\w+\s(C|c)hest[+-]\.definition"));
		}

		~ChestVoiceManager() {
			main.Dispose();
		}
		#endregion

		protected override void Control_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			string res = e.Result.Text;
			if (res == controlCommands.getStartCommand) {
				Console.Write("Chests opening mode initialized. Current type: ");
				if (main.Grammars.Count == 0) {
					Console.WriteLine("NOT SET!");
					Console.WriteLine("Set the chest type first with " + controlCommands.getSwitchGrammarCommand);
					return;
				}
				else {
					Console.WriteLine(main.Grammars[0].Name);
					main.Grammars[0].Enabled = true;
				}
			}
			else if (res == controlCommands.getStopCommand) {
				Console.WriteLine("Stopping Recognition!");
			}
			else if (res == controlCommands.getPauseCommand) {
				Console.WriteLine("Pausing Recognition!");
			}
			else if (res == controlCommands.getSwitchGrammarCommand) {
				Choices definitions = new Choices();
				Console.Write("Switching chest type, available: ");
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
			}
		}

		private void Switch_WordRecognized(object sender, SpeechRecognizedEventArgs e) {
			Console.WriteLine("\nSelected - " + e.Result.Text);
			main.UnloadAllGrammars();
			main.LoadGrammar(DefinitionParser.instance.GetGrammar(e.Result.Text));
			main.Grammars[0].Enabled = true;
			control.SpeechRecognized -= Switch_WordRecognized;
			control.Grammars[0].Enabled = true;
			control.SpeechRecognized += Control_SpeechRecognized;
			DefinitionParser.instance.currentGrammarFile = DefinitionParser.instance.GetDefinitionByName(e.Result.Text);
			DefinitionParser.instance.currentMobGrammarFile = null;
			Program.interaction.OpenWorksheet(e.Result.Text);
			Program.interaction.AutoAdjustColumns();
		}
	}
}
