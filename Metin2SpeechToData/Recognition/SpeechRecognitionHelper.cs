using System;
using System.Speech.Recognition;
using System.Windows.Forms;
using Metin2SpeechToData.Structures;
using System.Linq;
using System.Text;
using System.Collections.Generic;

namespace Metin2SpeechToData {
	public class SpeechRecognitionHelper : SpeechHelperBase {
		private enum UnderlyingRecognizer {
			AREA,
			CHEST,
			CUSTOM,
		}

		private readonly UnderlyingRecognizer currentMode;

		public GameRecognizer _gameRecognizer { get; }
		public ChestRecognizer _chestRecognizer { get; }

		#region Constructor
		public SpeechRecognitionHelper(GameRecognizer master) : base(master) {
			currentMode = UnderlyingRecognizer.AREA;
			InitializeControl();
			_gameRecognizer = master;
		}

		public SpeechRecognitionHelper(ChestRecognizer master) : base(master) {
			currentMode = UnderlyingRecognizer.CHEST;
			InitializeControl();
			_chestRecognizer = master;
		}

		/// <summary>
		/// Sets up controlling recognizer
		/// </summary>
		protected override void InitializeControl() {
			base.InitializeControl();
			Program.mapper.FreeControlHotkeys();
			Console.WriteLine("(F1){0} -- {1}\n" +
							  "(F4){2} -- Quit recognition",
							  CCommands.getSwitchGrammarCommand,
							  (currentMode != UnderlyingRecognizer.CHEST ? "Switch drop location" : "Change chest type"),
							  CCommands.getStopCommand);

			Program.mapper.AssignToHotkey(Keys.F1, Control_SpeechRecognized, new SpeechRecognizedArgs(CCommands.getSwitchGrammarCommand, 100));
			Program.mapper.AssignToHotkey(Keys.F4, Control_SpeechRecognized, new SpeechRecognizedArgs(CCommands.getStopCommand, 100));
		}
		#endregion


		/// <summary>
		/// Called by saying one of the controling words 
		/// </summary>
		protected override void Control_SpeechRecognized(SpeechRecognizedArgs e) {
			switch (e.textEnm) {
				case CCommands.Speech.START: {
					if (baseRecognizer.currentState == RecognitionBase.RecognitionState.ACTIVE) {
						Console.WriteLine("Already started!");
						break;
					}
					if (!baseRecognizer.isPrimaryDefinitionLoaded) {
						Console.WriteLine("Current grammar: NOT INITIALIZED!");
						Console.WriteLine("Set grammar first with " + CCommands.getSwitchGrammarCommand);
						break;
					}
					Console.Write("Starting Recognition...");
					Console.Write("Currently enabled words: ");

					StringBuilder list = new StringBuilder();
					foreach (string name in baseRecognizer.getCurrentGrammars.Keys) {
						list.Append(name + ", ");
					}
					Console.WriteLine(list.Remove(list.Length - 2, 2).ToString());
					baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.ACTIVE);
					Program.mapper.AttachHotkeyWrapper(_gameRecognizer);
					Program.mapper.FreeControlHotkeys();
					Program.mapper.FreeGameHotkeys();

					SetGrammarState(CCommands.getStartCommand, false);
					SetGrammarState(CCommands.getStopCommand, false);
					SetGrammarState(CCommands.getSwitchGrammarCommand, false);
					SetGrammarState(CCommands.getPauseCommand, true);
					SetGrammarState(CCommands.getDefineItemCommand, false);
					SetGrammarState(CCommands.getDefineMobCommand, false);

					Console.WriteLine("Pausing will reenable recognition control commands eg. " + CCommands.getStopCommand + ", " + CCommands.getSwitchGrammarCommand + "...");
					Console.WriteLine("To pause: " + KeyModifiers.Control + " + " + KeyModifiers.Shift + " + " + Keys.F4 +
									  " or '" + CCommands.getPauseCommand + "'");
					Program.mapper.AssignToHotkey(Keys.F4, KeyModifiers.Control, KeyModifiers.Shift, Control_SpeechRecognized,
												  new SpeechRecognizedArgs(CCommands.getPauseCommand, 100, true));
					Program.mapper.ToggleItemHotkeys(true);
					break;
				}
				case CCommands.Speech.STOP: {
					if (currentMode == UnderlyingRecognizer.AREA) {
						if (_gameRecognizer.enemyHandling.State == EnemyHandling.EnemyState.FIGHTING) {
							_gameRecognizer.enemyHandling.ForceKill();
						}
						Program.interaction.StopSession();
					}
					Program.mapper.DetachHotkeyWrapper();
					baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.STOPPED);
					baseRecognizer.Dispose();
					Dispose();
					ReturnControl();
					break;
				}
				case CCommands.Speech.PAUSE: {
					if (baseRecognizer.currentState != RecognitionBase.RecognitionState.ACTIVE) {
						if (e.asHotkey) {
							Console.WriteLine("Resuming recognition...");
							baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.ACTIVE);
							SetGrammarState(CCommands.getStartCommand, false);
							SetGrammarState(CCommands.getStopCommand, false);
							SetGrammarState(CCommands.getSwitchGrammarCommand, false);
							SetGrammarState(CCommands.getPauseCommand, true);
							SetGrammarState(CCommands.getDefineItemCommand, false);
							SetGrammarState(CCommands.getDefineMobCommand, false);
							Program.mapper.ToggleItemHotkeys(true);
							break;
						}
						Console.WriteLine("Recognition is not currenty active, no actions taken.");
						break;
					}
					baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.PAUSED);
					SetGrammarState(CCommands.getPauseCommand, false);
					SetGrammarState(CCommands.getStopCommand, true);
					SetGrammarState(CCommands.getStartCommand, true);
					SetGrammarState(CCommands.getDefineItemCommand, true);
					SetGrammarState(CCommands.getDefineMobCommand, true);
					Program.mapper.ToggleItemHotkeys(false);
					break;
				}
				case CCommands.Speech.SWITCH_GRAMMAR: {
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

					controlingRecognizer.LoadGrammar(new Grammar(definitions) { Name = "Available" });
					for (int i = 0; i < _currentGrammars.Count; i++) {
						controlingRecognizer.Grammars[i].Enabled = false;
						_currentGrammars[controlingRecognizer.Grammars[i].Name] = (i, false);
					}
					_currentGrammars.Add("Available", (_currentGrammars.Count, true));
					controlingRecognizer.SpeechRecognized += Switch_WordRecognized_Wrapper;
					controlingRecognizer.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
					baseRecognizer.OnRecognitionStateChanged(this, RecognitionBase.RecognitionState.SWITCHING);
					break;
				}
				case CCommands.Speech.DEFINE_MOB: {
					Undo.instance.currentOperationType = Undo.OperationTypes.Defining;
					Program.mapper.ToggleItemHotkeys(false);
					Console.WriteLine("Starting enemy definition creator");
					Console.Write("Write the main pronounciation: ");
					string mainPronoun = Console.ReadLine();
					if (!Undo.instance.ContinueExecution) {
						break;
					}
					Console.Write("Write ambiguities(alternate names) seperated by '/' or leave empty: ");
					string ambiguities = Console.ReadLine();
					if (!Undo.instance.ContinueExecution) {
						break;
					}
					ambiguities = ambiguities.Trim();
					Console.Write("Define mob level: ");
					string valString = Console.ReadLine();
					if (!Undo.instance.ContinueExecution) {
						break;
					}
					if (!ushort.TryParse(valString, out ushort level)) {
						Console.WriteLine("Entered invalid input, definition cancelled.");
						break;
					}
					Console.WriteLine("Choose mob group from: COMMON, HALF_BOSS, BOSS, METEOR, SPECIAL");
					string group = Console.ReadLine();
					if (!Undo.instance.ContinueExecution) {
						break;
					}
					group = group.ToUpper();
					if (group != "COMMON" && group != "HALF_BOSS" && group != "BOSS" && group != "METEOR" && group != "SPECIAL") {
						Console.WriteLine("Unknown group, cancelling");
						break;
					}

					
					string outputString;
					if (ambiguities == "") {
						outputString = (mainPronoun + "," + level.ToString() + "," + group);
					}
					else {
						outputString = (mainPronoun + "/" + ambiguities + "," + level.ToString() + "," + group);
					}
					List<string> ambList = new List<string>(ambiguities.Split('/'));
					int i = ambList.Count - 1;
					while (i >= 0) {
						if (ambList[i].Trim() == "") {
							ambList.Remove(ambList[i]);
						}
						i--;
					}
					Console.WriteLine(outputString);
					DefinitionParser.instance.AddMobEntry(outputString);
					MobParserData.Enemy enemy = new MobParserData.Enemy(mainPronoun, ambList.ToArray(), level, DefinitionParser.instance.currentMobGrammarFile.ParseClass(group));
					DefinitionParser.instance.currentMobGrammarFile.AddMobDuringRuntime(enemy);
					_gameRecognizer.enemyHandling.SwitchGrammar(DefinitionParser.instance.currentMobGrammarFile.ID.Split('_')[1]);

					Program.mapper.ToggleItemHotkeys(true);
					break;
				}
				case CCommands.Speech.DEFINE_ITEM: {
					Undo.instance.currentOperationType= Undo.OperationTypes.Defining;
					Program.mapper.ToggleItemHotkeys(false);
					Console.WriteLine("Starting item definition creator");
					Console.Write("Write the main pronounciation: ");
					string mainPronoun = Console.ReadLine();
					if (!Undo.instance.ContinueExecution) {
						break;
					}
					Console.Write("Write ambiguities(alternate names) seperated by '/' or leave empty: ");
					string ambiguities = Console.ReadLine();
					if (!Undo.instance.ContinueExecution) {
						break;
					}
					ambiguities = ambiguities.Trim();
					Console.Write("Define item value(as a positive value): ");
					
					string valString = Console.ReadLine();
					if (!Undo.instance.ContinueExecution) {
						break;
					}
					if (!uint.TryParse(valString, out uint value)) {
						Console.WriteLine("Entered invalid input, definition cancelled.");
						break;
					}
					Console.Write("Choose item group (group name): ");
					string group = Console.ReadLine();
					if (!Undo.instance.ContinueExecution) {
						break;
					}
					string outputString;
					if (ambiguities == "") {
						outputString = (mainPronoun + "," + value.ToString() + "," + group);
					}
					else {
						outputString = (mainPronoun + "/" + ambiguities + "," + value.ToString() + "," + group);
					}
					Console.WriteLine(outputString);
					DefinitionParser.instance.AddItemEntry(outputString, group);
					List<string> ambList = new List<string>(ambiguities.Split('/'));
					int i = ambList.Count - 1;
					while (i >= 0) {
						if (ambList[i].Trim() == "") {
							ambList.Remove(ambList[i]);
						}
						i--;
					}
					DefinitionParserData.Item item = new DefinitionParserData.Item(mainPronoun, ambList.ToArray(), value, group);
					DefinitionParser.instance.currentGrammarFile.AddItemDuringRuntime(item);
					_gameRecognizer.SwitchGrammar(DefinitionParser.instance.currentGrammarFile.ID);
					Program.mapper.ToggleItemHotkeys(true);
					break;
				}
				default: {
					Console.WriteLine("Undefined word " + e.text);
					break;
				}
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
					return DefinitionParser.instance.getDefinitionNames.Where(
						   (x) => new System.Text.RegularExpressions.Regex(@"C_\w+").IsMatch(x)).ToArray();
				}
				default: {
					return new string[0];
				}
			}
		}

		/// <summary>
		/// Called after saying 'switch command' followed by choice
		/// </summary>
		protected override void Switch_WordRecognized(SpeechRecognizedArgs e) {
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

			//TODO: Reimplement these cursed calls
			//controlingRecognizer.RecognizeAsyncCancel();
			controlingRecognizer.UnloadGrammar(controlingRecognizer.Grammars[_currentGrammars["Available"].index]);
			//controlingRecognizer.RecognizeAsync();

			Program.interaction.StartSession(e.text);
			Program.mapper.RemapHotkey(Keys.F1, Control_SpeechRecognized, new SpeechRecognizedArgs(CCommands.getStartCommand, 100));
			Program.mapper.AssignToHotkey(Keys.F2, Control_SpeechRecognized, new SpeechRecognizedArgs(CCommands.getStartSessionCommand, 100));
			Program.mapper.SetInactive(Keys.F4, false);

			Console.Clear();

			Console.WriteLine("Grammar initialized!");
			DefinitionParser.instance.LoadHotkeys(e.text);
			Console.WriteLine("(F1) or '" + CCommands.getStartCommand + "' to start\n" +
							  "(F4) or '" + CCommands.getStopCommand + "' to stop");
			baseRecognizer.currentState = RecognitionBase.RecognitionState.GRAMMAR_SELECTED;
		}
	}
}
