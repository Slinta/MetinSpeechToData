using System;
using System.Linq;
using System.Speech.Recognition;
using System.Windows.Forms;

namespace Metin2SpeechToData.Chests {
	class ChestVoiceManager : SpeechRecognitionHelper, IDisposable {

		private SpeechRecognitionEngine game;
		private DefinitionParser parser;
		private ChestSpeechRecognized recognition;

		#region Constructor / Destructor
		public ChestVoiceManager(ref SpeechRecognitionEngine engine) {
			game = engine;
			parser = new DefinitionParser(new System.Text.RegularExpressions.Regex(@"\w+\s(C|c)hest[+-]\.definition"));
			recognition = new ChestSpeechRecognized(this);
		}

		~ChestVoiceManager() {
			parser = null;
			game.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
		}
		#endregion

		protected override void Control_SpeechRecognized(SpeechRecognizedArgs e) {
			if (e.text == Program.controlCommands.getStartCommand) {
				Console.Write("Chests opening mode initialized. Current type: ");
				if (game.Grammars.Count == 0) {
					Console.WriteLine("NOT SET!");
					Console.WriteLine("Set the chest type first with " + Program.controlCommands.getSwitchGrammarCommand);
					return;
				}
				else {
					Console.WriteLine(game.Grammars[0].Name);
					game.Grammars[0].Enabled = true;
				}
				recognition.Subscribe(game);
				game.SetInputToDefaultAudioDevice();
				game.RecognizeAsync(RecognizeMode.Multiple);
				Console.WriteLine("Recognizing!");
			}
			else if (e.text == Program.controlCommands.getStopCommand) {
				Console.WriteLine("Stopping Recognition!");
				game.RecognizeAsyncStop();
			}
			else if (e.text == Program.controlCommands.getPauseCommand) {
				Console.WriteLine("Pausing Recognition!");
				game.RecognizeAsyncStop();
			}
			else if (e.text == Program.controlCommands.getSwitchGrammarCommand) {
				Choices definitions = new Choices();
				Console.Write("Switching chest type, available: ");
				string[] available = DefinitionParser.instance.getDefinitionNames;
				for (int i = 0; i < available.Length; i++) {
					definitions.Add(available[i]);
					Program.mapper.AssignToHotkey((Keys.D1 + i), Switch_WordRecognized, new SpeechRecognizedArgs(available[i], 100));
					if (i == available.Length - 1) {
						Console.Write("(" + (i + 1) + ")" + available[i]);
					}
					else {
						Console.Write("(" + (i + 1) + ")" + available[i] + ", ");
					}
				}
				if (control.Grammars.Count == 1) {
					control.LoadGrammar(new Grammar(definitions));
				}
				control.Grammars[0].Enabled = false;
				Program.mapper.SetInactive(Keys.F1, true);
				Program.mapper.SetInactive(Keys.F2, true);
				Program.mapper.SetInactive(Keys.F3, true);
				Program.mapper.SetInactive(Keys.F4, true);
				control.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
				control.SpeechRecognized += Switch_WordRecognized_Wrapper;
			}
		}

		private void Switch_WordRecognized_Wrapper(object sender, SpeechRecognizedEventArgs e) {
			Switch_WordRecognized(new SpeechRecognizedArgs(e.Result.Text, e.Result.Confidence));
		}
		private void Switch_WordRecognized(SpeechRecognizedArgs e) {
			Console.WriteLine("\nSelected - " + e.text);
			for (int i = (int)Keys.D1; i < (int)Keys.D9; i++) {
				Program.mapper.Free((Keys)i, true);
			}
			game.UnloadAllGrammars();
			game.LoadGrammar(DefinitionParser.instance.GetGrammar(e.text));
			game.LoadGrammar(new Grammar(new Choices(modifierDict.Values.ToArray())));
			game.Grammars[0].Enabled = true;
			game.Grammars[1].Enabled = true;
			control.SpeechRecognized -= Switch_WordRecognized_Wrapper;
			control.Grammars[0].Enabled = true;
			control.SpeechRecognized += Control_SpeechRecognized_Wrapper;
			DefinitionParser.instance.currentGrammarFile = DefinitionParser.instance.GetDefinitionByName(e.text);
			DefinitionParser.instance.currentMobGrammarFile = null;
			Program.interaction.OpenWorksheet(e.text);
			Program.mapper.SetInactive(Keys.F1, false);
			Program.mapper.SetInactive(Keys.F2, false);
			Program.mapper.SetInactive(Keys.F3, false);
			Program.mapper.SetInactive(Keys.F4, false);
			Console.Clear();
			Console.WriteLine("Grammar initialized!");
			DefinitionParser.instance.LoadHotkeys(e.text);
			Console.WriteLine("(F1) or '" + Program.controlCommands.getStartCommand + "' to start");
		}

		#region IDisposable Support
		private bool disposedValue = false;

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				if (disposing) {
					control.Dispose();
					game.Dispose();
					recognition.Dispose();
				}

				disposedValue = true;
			}
		}

		// This code added to correctly implement the disposable pattern.
		public void Dispose() {
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
