using System;
using System.Linq;
using System.Collections.Generic;
using System.Speech.Recognition;
using System.IO;
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
			control.SpeechRecognized += Control_SpeechRecognized;
			if (Program.debug) {
				Console.WriteLine("Control grammar loaded...");
			}
			Console.Write("Available commands:\n" +
						  Program.controlCommands.getStartCommand + " - Start recognition\n" +
						  Program.controlCommands.getPauseCommand + " - Pauses main recognition\n" +
						  Program.controlCommands.getStopCommand + " - Exits App\n" +
						  Program.controlCommands.getSwitchGrammarCommand + " - Changes grammar (your drop location)\n");
			if (!Program.debug) {
				control.RecognizeAsync(RecognizeMode.Multiple);
			}
		}
		#endregion

		protected virtual void Control_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			string res = e.Result.Text;

			if (res == Program.controlCommands.getStartCommand) {
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
			else if (res == Program.controlCommands.getStopCommand) {
				Console.WriteLine("Stopping Recognition!");
			}
			else if (res == Program.controlCommands.getPauseCommand) {
				Console.WriteLine("Pausing Recognition!");
				OnRecognitionChange(Program.RecognitionState.CONTROL_RUNNING);

			}
			else if (res == Program.controlCommands.getSwitchGrammarCommand) {
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
			Program.interaction.OpenWorksheet(e.Result.Text);
			if (Program.debug) {
				Console.WriteLine(main.Grammars.Count);
			}
		}

		/// <summary>
		/// Prevents Console.ReadLine() from Main to consume lines typed for different prompt
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
			if (Program.debug) {
				Console.WriteLine("Waited long enough!");
			}
		}
	}
}
