using System;

namespace Metin2SpeechToData {
	public class RecognitionSetup {

		public RecognitionSetup() {
			Begin();
		}

		public void Begin() {

		}

		/// <summary>
		/// Sets up controlling recognizer
		/// </summary>
		private void InitializeControl() {
			controlingRecognizer = new SpeechRecognitionEngine();
			_currentGrammars = new Dictionary<string, (int, bool)>();

			//string startC = Program.controlCommands.getStartCommand;
			//string startSessionC = Program.controlCommands.getStartSessionCommand;
			//string pauseC = Program.controlCommands.getPauseCommand;
			string switchC = Program.controlCommands.getSwitchGrammarCommand;
			string quitC = Program.controlCommands.getStopCommand;

			//controlingRecognizer.LoadGrammar(new Grammar(new Choices(startC)));
			//controlingRecognizer.LoadGrammar(new Grammar(new Choices(startSessionC)));
			//controlingRecognizer.LoadGrammar(new Grammar(new Choices(pauseC)));
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(switchC)));
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(quitC)));

			//_currentGrammars.Add(startC, (0, true));
			//_currentGrammars.Add(startSessionC, (1, true));
			//_currentGrammars.Add(pauseC, (2, true));
			_currentGrammars.Add(switchC, (1, true));
			_currentGrammars.Add(quitC, (2, true));

			controlingRecognizer.SetInputToDefaultAudioDevice();
			controlingRecognizer.SpeechRecognized += Control_SpeechRecognized_Wrapper;

			if (Program.debug) {
				Console.WriteLine("Control grammar loaded...");
			}
			Program.mapper.FreeControlHotkeys();

			Console.Write("Available commands:\n" +
						  //startC + " - Start recognition(F1)\n" +
						  //startSessionC + " - Start recognition as session(F6)\n\n" +
						  //pauseC + " - Pauses main recognition(F2)\n" +
						  switchC + " - Changes grammar (your drop location)(F1)\n" +
						  quitC + " - Exits App(F4)\n");
			//Program.mapper.AssignToHotkey(Keys.F1, Control_SpeechRecognized, new SpeechRecognizedArgs(startC, 100));
			//Program.mapper.AssignToHotkey(Keys.F6, Control_SpeechRecognized, new SpeechRecognizedArgs(startSessionC, 100));
			//Program.mapper.AssignToHotkey(Keys.F2, Control_SpeechRecognized, new SpeechRecognizedArgs(pauseC, 100));
			Program.mapper.AssignToHotkey(Keys.F1, Control_SpeechRecognized, new SpeechRecognizedArgs(switchC, 100));
			Program.mapper.AssignToHotkey(Keys.F4, Control_SpeechRecognized, new SpeechRecognizedArgs(quitC, 100));

			controlingRecognizer.RecognizeAsync(RecognizeMode.Multiple);
		}

		/// <summary>
		/// Called by saying one of the controling words 
		/// </summary>
		private void Control_SpeechRecognized_Wrapper(object sender, SpeechRecognizedEventArgs e) {
			if (e.Result.Confidence > Configuration.acceptanceThreshold) {
				Control_SpeechRecognized(new SpeechRecognizedArgs(e.Result.Text, e.Result.Confidence));
			}
		}
		private void Control_SpeechRecognized(SpeechRecognizedArgs e) {
			if (e.text == Program.controlCommands.getStartCommand) {
				if (baseRecognizer.currentState == RecognitionBase.RecognitionState.ACTIVE) {
					Console.WriteLine("Already started!");
					return;
				}

				Console.Write("Starting Recognition... ");
				if (!baseRecognizer.isPrimaryDefinitionLoaded) {
					Console.WriteLine("Current grammar: NOT INITIALIZED!");
					Console.WriteLine("Set grammar first with " + Program.controlCommands.getSwitchGrammarCommand);
				}
				else {
					Console.Write("Currently enabled grammars: ");
					string list = "";
					foreach (string name in baseRecognizer.getCurrentGrammars.Keys) {
						list = list + name + ", ";
					}
					Console.WriteLine(list.Remove(list.Length - 2, 2));
					baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.ACTIVE);
					Program.mapper.FreeNonCustomHotkeys();

					SetGrammarActive(Program.controlCommands.getStartCommand, false);
					SetGrammarActive(Program.controlCommands.getStartSessionCommand, false);
					SetGrammarActive(Program.controlCommands.getStopCommand, false);
					SetGrammarActive(Program.controlCommands.getSwitchGrammarCommand, false);
					SetGrammarActive(Program.controlCommands.getPauseCommand, true);

					Console.WriteLine("To pause: " + KeyModifiers.Control + " + " + KeyModifiers.Shift + " + " + Keys.F4 + " or '" + Program.controlCommands.getPauseCommand + "'");
					Console.WriteLine("Pausing will enable the rest of control.");
					Program.mapper.AssignToHotkey(Keys.F4, KeyModifiers.Control, KeyModifiers.Shift, Control_SpeechRecognized, new SpeechRecognizedArgs(Program.controlCommands.getPauseCommand, 100));
					DefinitionParser.instance.hotkeyParser.SetKeysActiveState(true);
				}
			}
			else if (e.text == Program.controlCommands.getStartSessionCommand) {
				if (baseRecognizer.currentState == RecognitionBase.RecognitionState.ACTIVE) {
					Console.WriteLine("Already started!");
					return;
				}
				Console.Write("Starting Recognition... ");
				if (!baseRecognizer.isPrimaryDefinitionLoaded) {
					Console.WriteLine("Current grammar: NOT INITIALIZED!");
					Console.WriteLine("Set grammar first with " + Program.controlCommands.getSwitchGrammarCommand);
					return;
				}
				Program.interaction.SwitchFile(default(System.IO.FileInfo));
				Control_SpeechRecognized(new SpeechRecognizedArgs(Program.controlCommands.getStartCommand, 100));
			}
			else if (e.text == Program.controlCommands.getStopCommand) {
				ReturnControl();
				Console.WriteLine("Stopping Voice Recognition!");
			}
			else if (e.text == Program.controlCommands.getPauseCommand) {
				if (baseRecognizer.currentState != RecognitionBase.RecognitionState.ACTIVE) {
					Console.WriteLine("Recognition is not currenty active, no actions taken.");
					return;
				}

				baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.PAUSED);
				SetGrammarActive(Program.controlCommands.getPauseCommand, false);
				SetGrammarActive(Program.controlCommands.getStopCommand, true);
				SetGrammarActive(Program.controlCommands.getStartCommand, true);
				DefinitionParser.instance.hotkeyParser.SetKeysActiveState(false);
			}
			else if (e.text == Program.controlCommands.getSwitchGrammarCommand) {
				if (baseRecognizer.currentState >= RecognitionBase.RecognitionState.GRAMMAR_SELECTED) {
					Console.WriteLine("You can not select another grammar at this point, Pause >> Quit >> Select grammar");
					return;
				}
				Choices definitions = new Choices();
				Console.Write("Switching Grammar, available: ");
				string[] available = DefinitionParser.instance.getDefinitionNames;
				for (int i = 0; i < available.Length; i++) {
					definitions.Add(available[i]);
					Program.mapper.SetInactive(Keys.F1, true);
					Program.mapper.SetInactive(Keys.F2, true);
					//Program.mapper.SetInactive(Keys.F3, true);
					//Program.mapper.SetInactive(Keys.F4, true);
					Program.mapper.AssignToHotkey((Keys.D1 + i), Switch_WordRecognized, new SpeechRecognizedArgs(available[i], 100));
					if (i == available.Length - 1) {
						Console.Write("(" + (i + 1) + ")" + available[i]);
					}
					else {
						Console.Write("(" + (i + 1) + ")" + available[i] + ", ");
					}
				}

				controlingRecognizer.LoadGrammar(new Grammar(definitions));
				for (int i = 0; i < _currentGrammars.Count; i++) {
					controlingRecognizer.Grammars[i].Enabled = false;
				}
				controlingRecognizer.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
				controlingRecognizer.SpeechRecognized += Switch_WordRecognized_Wrapper;
				baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.SWITCHING);
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
			}
			controlingRecognizer.UnloadGrammar(controlingRecognizer.Grammars[_currentGrammars.Count]);

			Program.interaction.OpenWorksheet(e.text);
			Program.mapper.FreeSpecific(Keys.F1, true);
			Program.mapper.AssignToHotkey(Keys.F1, Control_SpeechRecognized, new SpeechRecognizedArgs(Program.controlCommands.getStartCommand, 100));
			Program.mapper.AssignToHotkey(Keys.F2, Control_SpeechRecognized, new SpeechRecognizedArgs(Program.controlCommands.getStartSessionCommand, 100));
			//Program.mapper.SetInactive(Keys.F1, false);
			//Program.mapper.SetInactive(Keys.F2, true);
			//Program.mapper.SetInactive(Keys.F3, true);
			Program.mapper.SetInactive(Keys.F4, false);

			Console.Clear();
			Console.WriteLine("Grammar initialized!");
			DefinitionParser.instance.LoadHotkeys(e.text);
			Console.WriteLine("(F1) or '" + Program.controlCommands.getStartCommand + "' to start\n" +
							  "(F4) or '" + Program.controlCommands.getStopCommand + "' to stop");
			baseRecognizer.currentState = RecognitionBase.RecognitionState.GRAMMAR_SELECTED;
		}


		private readonly ManualResetEventSlim signal = new ManualResetEventSlim();

		/// <summary>
		/// Prevents Console.ReadLine() from Main from consuming lines meant for different prompt
		/// </summary>
		public void AcquireControl() {
			EventHandler<SpeechRecognizedEventArgs> waitCancellation = new EventHandler<SpeechRecognizedEventArgs>(
			(object o, SpeechRecognizedEventArgs e) => {
				if (e.Result.Text == Program.controlCommands.getStopCommand) {
					signal.Set();
				}
			});

			controlingRecognizer.SpeechRecognized += waitCancellation;
			signal.Wait();
			controlingRecognizer.RecognizeAsyncStop();
			controlingRecognizer.SpeechRecognized -= waitCancellation;
			controlingRecognizer.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
			controlingRecognizer.Dispose();
		}

		private void ReturnControl() {
			signal.Set();
		}
	}
}
