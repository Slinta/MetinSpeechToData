using System;
using System.Collections.Generic;
using System.Speech.Recognition;
using System.Windows.Forms;
using Metin2SpeechToData.Structures;
using System.Linq;
using System.Text;

namespace Metin2SpeechToData {
	public class SpeechRecognitionHelper : SpeechHelperBase {

		private enum UnderlyingRecognizer {
			AREA,
			CHEST,
			CUSTOM,
		}

		private UnderlyingRecognizer currentMode;

		#region Constructor
		public SpeechRecognitionHelper(GameRecognizer master): base(master) {
			currentMode = UnderlyingRecognizer.AREA;
			InitializeControl();
		}

		public SpeechRecognitionHelper(ChestRecognizer master) : base(master) {
			currentMode = UnderlyingRecognizer.CHEST;
			InitializeControl();
		}

		/// <summary>
		/// Sets up controlling recognizer
		/// </summary>
		protected override void InitializeControl() {
			base.InitializeControl();

			Program.mapper.FreeControlHotkeys();
			Console.Write("Available commands:\n" +
						  CCommands.getSwitchGrammarCommand + " - Changes grammar (your drop location)(F1)\n" +
						  CCommands.getStopCommand + " - Exits App(F4)\n");

			Program.mapper.AssignToHotkey(Keys.F1, Control_SpeechRecognized, new SpeechRecognizedArgs(CCommands.getSwitchGrammarCommand, 100));
			Program.mapper.AssignToHotkey(Keys.F4, Control_SpeechRecognized, new SpeechRecognizedArgs(CCommands.getStopCommand, 100));
		}
		#endregion


		/// <summary>
		/// Called by saying one of the controling words 
		/// </summary>
		protected override void Control_SpeechRecognized(SpeechRecognizedArgs e) {

			if (e.text == CCommands.getStartCommand) {
				if (baseRecognizer.currentState == RecognitionBase.RecognitionState.ACTIVE) {
					Console.WriteLine("Already started!");
					return;
				}

				Console.Write("Starting Recognition... ");
				if (!baseRecognizer.isPrimaryDefinitionLoaded) {
					Console.WriteLine("Current grammar: NOT INITIALIZED!");
					Console.WriteLine("Set grammar first with " + CCommands.getSwitchGrammarCommand);
				}
				else {
					Console.Write("Currently enabled grammars: ");
					StringBuilder list = new StringBuilder();
					foreach (string name in baseRecognizer.getCurrentGrammars.Keys) {
						list.Append(name + ", ");
					}
					Console.WriteLine(list.Remove(list.Length - 2, 2).ToString());
					baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.ACTIVE);
					Program.mapper.FreeNonCustomHotkeys();

					SetGrammarState(CCommands.getStartCommand, false);
					SetGrammarState(CCommands.getStopCommand, false);
					SetGrammarState(CCommands.getSwitchGrammarCommand, false);
					SetGrammarState(CCommands.getPauseCommand, true);

					Console.WriteLine("To pause: " + KeyModifiers.Control + " + " + KeyModifiers.Shift + " + " + Keys.F4 + " or '" + CCommands.getPauseCommand + "'");
					Console.WriteLine("Pausing will enable the rest of control.");
					Program.mapper.AssignToHotkey(Keys.F4, KeyModifiers.Control, KeyModifiers.Shift, Control_SpeechRecognized,
												  new SpeechRecognizedArgs(CCommands.getPauseCommand, 100, true));
					DefinitionParser.instance.hotkeyParser.SetKeysActiveState(true);
				}
			}
			else if (e.text == CCommands.getStopCommand) {
				baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.STOPPED);
				baseRecognizer.Dispose();
				Dispose();
				ReturnControl();
			}
			else if (e.text == CCommands.getPauseCommand) {
				if (baseRecognizer.currentState != RecognitionBase.RecognitionState.ACTIVE) {
					if (e.asHotkey) {
						Console.WriteLine("Resuming recognition...");
						baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.ACTIVE);
						SetGrammarState(CCommands.getStartCommand, false);
						SetGrammarState(CCommands.getStopCommand, false);
						SetGrammarState(CCommands.getSwitchGrammarCommand, false);
						SetGrammarState(CCommands.getPauseCommand, true);
						DefinitionParser.instance.hotkeyParser.SetKeysActiveState(true);
						return;
					}
					Console.WriteLine("Recognition is not currenty active, no actions taken.");
					return;
				}

				baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.PAUSED);
				SetGrammarState(CCommands.getPauseCommand, false);
				SetGrammarState(CCommands.getStopCommand, true);
				SetGrammarState(CCommands.getStartCommand, true);
				DefinitionParser.instance.hotkeyParser.SetKeysActiveState(false);
			}
			else if (e.text == CCommands.getSwitchGrammarCommand) {
				if (baseRecognizer.currentState >= RecognitionBase.RecognitionState.GRAMMAR_SELECTED) {
					Console.WriteLine("You can not select another grammar at this point, Pause >> Quit >> Select grammar");
					return;
				}
				Choices definitions = new Choices();
				Console.Write("Switching Grammar, available: ");
				string[] available = GetDefinitionNames(currentMode);
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

				Program.mapper.SetInactive(Keys.F1, true);
				Program.mapper.SetInactive(Keys.F2, true);
				Program.mapper.SetInactive(Keys.F3, true);
				Program.mapper.SetInactive(Keys.F4, true);

				controlingRecognizer.LoadGrammar(new Grammar(definitions));
				for (int i = 0; i < _currentGrammars.Count; i++) {
					controlingRecognizer.Grammars[i].Enabled = false;
					_currentGrammars[controlingRecognizer.Grammars[i].Name] = (i, false);
				}
				controlingRecognizer.SpeechRecognized += Switch_WordRecognized_Wrapper;
				controlingRecognizer.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
				baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.SWITCHING);
			}
		}

		private string[] GetDefinitionNames(UnderlyingRecognizer mode) {
			switch (mode) {
				case UnderlyingRecognizer.AREA: {
					return DefinitionParser.instance.getDefinitionNames.Where(
						   (x) => new System.Text.RegularExpressions.Regex(Configuration.AREA_REGEXP).IsMatch(x)).ToArray();
				}
				case UnderlyingRecognizer.CHEST: {
					return DefinitionParser.instance.getDefinitionNames.Where(
						   (x) => new System.Text.RegularExpressions.Regex(Configuration.CHESTS_REGEXP).IsMatch(x)).ToArray();
				}
				case UnderlyingRecognizer.CUSTOM: {
					//TODO
					throw new NotImplementedException();
				}
				default: {
					return new string[0];
				}
			}
		}

		/// <summary>
		/// Called whenever user says SWITCH_COMMAND, which causes controling recognizer to switch to this method
		/// to handle the next word that will signalize the location of choice
		/// </summary>
		private void Switch_WordRecognized_Wrapper(object sender, SpeechRecognizedEventArgs e) {
			if (Configuration.acceptanceThreshold < e.Result.Confidence) {
				Switch_WordRecognized(new SpeechRecognizedArgs(e.Result.Text, e.Result.Confidence));
			}
		}
		private void Switch_WordRecognized(SpeechRecognizedArgs e) {

			controlingRecognizer.SpeechRecognized -= Switch_WordRecognized_Wrapper;
			controlingRecognizer.SpeechRecognized += Control_SpeechRecognized_Wrapper;

			Console.WriteLine("\nSelected - " + e.text);
			for (int i = (int)Keys.D1; i < (int)Keys.D9; i++) {
				Program.mapper.FreeSpecific((Keys)i, true);
			}

			DefinitionParser.instance.UpdateCurrents(e.text);

			baseRecognizer.SwitchGrammar(e.text);
			baseRecognizer.isPrimaryDefinitionLoaded = true;


			for (int i = 0; i < _currentGrammars.Count; i++) {
				controlingRecognizer.Grammars[i].Enabled = true;
				_currentGrammars[controlingRecognizer.Grammars[i].Name] = (i, true);
			}
			controlingRecognizer.UnloadGrammar(controlingRecognizer.Grammars[_currentGrammars.Count]);

			Program.interaction.OpenWorksheet(e.text);
			Program.mapper.RemapHotkey(Keys.F1, Control_SpeechRecognized, new SpeechRecognizedArgs(CCommands.getStartCommand, 100));
			Program.mapper.AssignToHotkey(Keys.F2, Control_SpeechRecognized, new SpeechRecognizedArgs(CCommands.getStartSessionCommand, 100));
			Program.mapper.SetInactive(Keys.F4, false);

			Console.Clear();

			Console.WriteLine("Grammar initialized!");
			DefinitionParser.instance.LoadHotkeys(e.text);
			Console.WriteLine("(F1) or '" + CCommands.getStartCommand + "' to start\n" +
							  "(F2) or '" + CCommands.getStartSessionCommand + "'to start as session\n" +
							  "(F4) or '" + CCommands.getStopCommand + "' to stop");
			baseRecognizer.currentState = RecognitionBase.RecognitionState.GRAMMAR_SELECTED;
		}
	}

	public sealed class ModiferRecognizedEventArgs : EventArgs {
		public CCommands.Speech modifier { get; set; }
		public string triggeringItem { get; set; }
	}
}
