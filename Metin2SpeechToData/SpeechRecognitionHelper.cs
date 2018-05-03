using System;
using System.Linq;
using System.Collections.Generic;
using System.Speech.Recognition;
using System.Windows.Forms;
using System.Threading;

namespace Metin2SpeechToData {
	public class SpeechRecognitionHelper {

		public enum ModifierWords {
			NONE,
			NEW_TARGET,
			UNDO,
			TARGET_KILLED,
		};

		/// <summary>
		/// Modifiers dictionary, used to convert enum values to spoken word
		/// </summary>
		public static IReadOnlyDictionary<ModifierWords, string> modifierDict = new Dictionary<ModifierWords, string>() {
			{ ModifierWords.NEW_TARGET , Program.controlCommands.getNewTargetCommand },
			{ ModifierWords.TARGET_KILLED, Program.controlCommands.getTargetKilledCommand },
			{ ModifierWords.UNDO, Program.controlCommands.getUndoCommand },
		};
		public static IReadOnlyDictionary<string, ModifierWords> reverseModifierDict = new Dictionary<string, ModifierWords>() {
			{ Program.controlCommands.getNewTargetCommand , ModifierWords.NEW_TARGET },
			{ Program.controlCommands.getTargetKilledCommand, ModifierWords.TARGET_KILLED },
			{ Program.controlCommands.getUndoCommand, ModifierWords.UNDO },
		};

		protected SpeechRecognitionEngine control;
		private SpeechRecognitionEngine main;

		public static ModifierWords currentModifier = ModifierWords.NONE;

		public event Program.Recognition OnRecognitionChange;

		#region Contrusctor / Destructor
		public SpeechRecognitionHelper() {
			main = null;
			InitializeControl();
		}

		public SpeechRecognitionHelper(ref SpeechRecognitionEngine engine) {
			main = engine;
			InitializeControl();
		}

		~SpeechRecognitionHelper() {
			Console.WriteLine("Destructor of SpeechRecognitionHelper.");
		}

		public void CleanUp() {
			currentModifier = ModifierWords.NONE;
			control.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
			control.SpeechRecognized -= Switch_WordRecognized_Wrapper;

		}

		private void InitializeControl() {
			control = new SpeechRecognitionEngine();
			Grammar controlGrammar = new Grammar(new Choices(
					Program.controlCommands.getStartCommand,
					Program.controlCommands.getPauseCommand,
					Program.controlCommands.getStopCommand,
					Program.controlCommands.getSwitchGrammarCommand)) {
				Name = "Controler Grammar"
			};

			control.LoadGrammar(controlGrammar);
			control.SetInputToDefaultAudioDevice();
			control.SpeechRecognized += Control_SpeechRecognized_Wrapper;
			if (Program.debug) {
				Console.WriteLine("Control grammar loaded...");
			}
			Program.mapper.FreeControl();

			Console.Write("Available commands:\n" +
						  Program.controlCommands.getStartCommand + " - Start recognition(F1)\n" +
						  Program.controlCommands.getPauseCommand + " - Pauses main recognition(F2)\n" +
						  Program.controlCommands.getSwitchGrammarCommand + " - Changes grammar (your drop location)(F3)\n" +
						  Program.controlCommands.getStopCommand + " - Exits App(F4)\n");
			Program.mapper.AssignToHotkey(Keys.F1, Control_SpeechRecognized, new SpeechRecognizedArgs(Program.controlCommands.getStartCommand, 100));
			Program.mapper.AssignToHotkey(Keys.F2, Control_SpeechRecognized, new SpeechRecognizedArgs(Program.controlCommands.getPauseCommand, 100));
			Program.mapper.AssignToHotkey(Keys.F3, Control_SpeechRecognized, new SpeechRecognizedArgs(Program.controlCommands.getSwitchGrammarCommand, 100));
			Program.mapper.AssignToHotkey(Keys.F4, Control_SpeechRecognized, new SpeechRecognizedArgs(Program.controlCommands.getStopCommand, 100));

			if (!Program.debug) {
				control.RecognizeAsync(RecognizeMode.Multiple);
			}
		}
		#endregion

		protected void Control_SpeechRecognized_Wrapper(object sender, SpeechRecognizedEventArgs e) {
			Control_SpeechRecognized(new SpeechRecognizedArgs(e.Result.Text, e.Result.Confidence));
		}
		protected virtual void Control_SpeechRecognized(SpeechRecognizedArgs e) {

			if (e.text == Program.controlCommands.getStartCommand) {
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
					Program.mapper.FreeAll();
					currentModifier = ModifierWords.NONE;

					LoadNewControlGrammar(MenuGrammarWithout(new string[1] { Program.controlCommands.getStartCommand }));
				}
			}
			else if (e.text == Program.controlCommands.getStopCommand) {
				Console.WriteLine("Stopping Recognition!");
			}
			else if (e.text == Program.controlCommands.getPauseCommand) {
				Console.WriteLine("Pausing Recognition!");
				LoadNewControlGrammar(MenuGrammarWithout(new string[0]));
				OnRecognitionChange(Program.RecognitionState.CONTROL_RUNNING);

			}
			else if (e.text == Program.controlCommands.getSwitchGrammarCommand) {
				Choices definitions = new Choices();
				Console.Write("Switching Grammar, available: ");
				string[] available = DefinitionParser.instance.getDefinitionNames;
				for (int i = 0; i < available.Length; i++) {
					definitions.Add(available[i]);
					Program.mapper.SetInactive(Keys.F1, true);
					Program.mapper.SetInactive(Keys.F2, true);
					Program.mapper.SetInactive(Keys.F3, true);
					Program.mapper.SetInactive(Keys.F4, true);
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
				control.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
				control.SpeechRecognized += Switch_WordRecognized_Wrapper;
				OnRecognitionChange(Program.RecognitionState.SWITCHING);
			}
			else {
				OnRecognitionChange(Program.RecognitionState.ERROR);
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
			main.UnloadAllGrammars();
			main.LoadGrammar(DefinitionParser.instance.GetGrammar(e.text));
			main.LoadGrammar(new Grammar(new Choices(modifierDict.Values.ToArray())));
			main.LoadGrammar(DefinitionParser.instance.GetMobGrammar("Mob_" + e.text));
			main.Grammars[0].Enabled = true;
			main.Grammars[1].Enabled = true;
			main.Grammars[2].Enabled = false;
			control.SpeechRecognized -= Switch_WordRecognized_Wrapper;
			control.Grammars[0].Enabled = true;
			control.SpeechRecognized += Control_SpeechRecognized_Wrapper;
			control.Grammars[1].Enabled = false;
			DefinitionParser.instance.currentGrammarFile = DefinitionParser.instance.GetDefinitionByName(e.text);
			DefinitionParser.instance.currentMobGrammarFile = DefinitionParser.instance.GetMobDefinitionByName("Mob_" + e.text);
			Program.interaction.OpenWorksheet(e.text);
			Program.mapper.SetInactive(Keys.F1, false);
			Program.mapper.SetInactive(Keys.F2, false);
			Program.mapper.SetInactive(Keys.F3, false);
			Program.mapper.SetInactive(Keys.F4, false);
			if (Program.debug) {
				Console.WriteLine(main.Grammars.Count);
			}
		}

		/// <summary>
		/// Prevents Console.ReadLine() from Main from consuming lines meant for different prompt
		/// </summary>
		public void AcquireControl() {
			ManualResetEventSlim signal = new ManualResetEventSlim();
			EventHandler<SpeechRecognizedEventArgs> handle = new EventHandler<SpeechRecognizedEventArgs>(
				(object o, SpeechRecognizedEventArgs e) => {
					if (e.Result.Text == Program.controlCommands.getStopCommand) {
						signal.Set();
					}
				}
			);
			control.SpeechRecognized += handle;
			signal.Wait();
			//TODO clean references here more ??
			control.RecognizeAsyncStop();
			control.Dispose();
			control.SpeechRecognized -= handle;
			control.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
			
			if (Program.debug) {
				Console.WriteLine("Waited long enough!");
			}
		}


		private void LoadNewControlGrammar(Grammar grammar) {
			control.UnloadGrammar(control.Grammars[0]);
			control.LoadGrammar(grammar);
		}

		public Grammar MenuGrammarWithout(string[] forbiddenC) {
			Choices choices = new Choices();
			IEnumerable<string> resultingCommands = Program.controlCommands.MenuCommands().Except(forbiddenC);
			choices.Add(resultingCommands.ToArray());
			Grammar grammar = ( new Grammar(choices) {
				Name = "Controler Grammar"
			});
			return grammar;

		}
		public Grammar MenuGrammarWithout(List<string> forbiddenC) {
			Choices choices = new Choices();
			IEnumerable<string> resultingCommands = Program.controlCommands.MenuCommands().Except(forbiddenC);
			choices.Add(resultingCommands.ToArray());
			Grammar grammar = (new Grammar(choices) {
				Name = "Controler Grammar"
			});
			return grammar;

		}
	}
}
