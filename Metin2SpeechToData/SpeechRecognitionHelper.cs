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
			REMOVE_TARGET,
			UNDO,
			TARGET_KILLED
		};

		/// <summary>
		/// Modifiers dictionary, used to convert enum values to spoken word
		/// </summary>
		public static IReadOnlyDictionary<ModifierWords, string> modifierDict = new Dictionary<ModifierWords, string>() {
			{ ModifierWords.NEW_TARGET , "New Target" },
			{ ModifierWords.TARGET_KILLED, "Killed Target" },
			{ ModifierWords.REMOVE_TARGET, "Remove Target" },
			{ ModifierWords.UNDO, "Undo" },
		};

		protected SpeechRecognitionEngine control;
		private SpeechRecognitionEngine main;

		public ControlSpeechCommands controlCommands { get; private set; }

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
			if (!Program.debug) {
				control.RecognizeAsync(RecognizeMode.Multiple);
			}
		}
		#endregion

		protected virtual void Control_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
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
					if (e.Result.Text == controlCommands.getStopCommand) {
						signal.Set();
					}
				}
			);
			control.SpeechRecognized += handle;
			signal.Wait();
			//TODO clean references here more ??
			main.RecognizeAsyncStop();
			control.RecognizeAsyncStop();
			main.Dispose();
			control.Dispose();
			control.SpeechRecognized -= handle;
			if (Program.debug) {
				Console.WriteLine("Waited long enough!");
			}
		}
	}
}
