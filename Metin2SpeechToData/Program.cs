using System;
using System.IO;
using System.Speech.Recognition;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public class Program {
		public enum RecognitionState {
			ERROR,
			CONTROL_RUNNING,
			DROP_LOGGER_RUNNING,
			SWITCHING,
		}

		public delegate void Recognition(RecognitionState state);
		public delegate void ModifierTrigger(SpeechRecognitionHelper.ModifierWords word, params string[] args);

		public static event ModifierTrigger OnModifierWordHear;

		public static bool debug = false;

		private static SpeechRecognitionEngine game;
		public static SpeechRecognitionHelper helper;

		private static DefinitionParser parser;

		public static EnemyHandling enemyHandling;
		public static SpreadsheetInteraction interaction;

		public static Configuration config;
		public static ControlSpeechCommands controlCommands;

		public static HotKeyMapper mapper;

		public static string currCommand = "";


		/* NIY List:
		 * TODO: Sessions
		 *
		 * 
		 * 
		 */

		[STAThread]
		static void Main(string[] args) {
			// Init
			config = new Configuration(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg");
			interaction = new SpreadsheetInteraction(config.xlsxFile);
			controlCommands = new ControlSpeechCommands("Control.definition");
			Confirmation.Initialize();

			mapper = new HotKeyMapper();
			mapper.AssignToHotkey(Keys.F1, "voice");
			mapper.AssignToHotkey(Keys.F2, "chest");
			mapper.AssignToHotkey(Keys.F3, "help");
			mapper.AssignToHotkey(Keys.F4, "quit");
			mapper.AssignToHotkey(Keys.F8, KeyModifiers.Shift, "wipe");

			bool continueRunning = true;
			//

			Console.WriteLine("Welcome to Metin2 siNDiCATE Drop logger");
			Console.WriteLine("Type 'help' for more info on how to use this program");

			while (continueRunning) {
				Console.WriteLine("Commands:" +
					"\n-(F1) Voice recognition" +
					"\n-(F2) Chests" +
					"\n-(F3) Help" +
					"\n-(F4) Quit" +
					"\n-(Shift + F8) Wipe");

				string command = Console.ReadLine();
				if (currCommand != "") {
					command = currCommand;
				}

				string[] commandBlocks = command.Split(' ');

				//Switch over length
				switch (commandBlocks.Length) {
					case 1: {
						switch (commandBlocks[0]) {
							case "quit": {
								Console.WriteLine("Do you want to quit? y/n");
								if (Console.ReadKey().Key == ConsoleKey.Y) {
									continueRunning = false;
								}
								break;
							}
							case "exit": {
								Console.WriteLine("Do you want to quit? y/n");
								if (Console.ReadKey().Key == ConsoleKey.Y) {
									continueRunning = false;
								}
								break;
							}
							case "help": {
								//TODO Update help print
								Console.WriteLine("Existing commands:");
								Console.WriteLine("quit / exit --> Close the application\n");
								Console.WriteLine("clear --> Clears the console\n");
								Console.WriteLine("voice / voice debug --> Enables voice control without/with debug prints\n");
								Console.WriteLine("chest --> Opens speech recognition for chest drops (Gold/Silver[+-])\n");
								Console.WriteLine("sheet 'sheet name' --> Swithches current working sheet to 'sheet name'" +
												  "\n'sheet name' = sheet name, sheet MUST already exist!\n");
								Console.WriteLine("sheet add 'name' --> adds a sheet with name 'name'\n");
								Console.WriteLine("[Deprecated for single item addition]\n" +
												  "add 'row' 'collumn' 'number' --> Adds value to cell\n" +
												  "'number' = the number to add\n" +
												  "'row', 'collumn' = indexes of the cell\n" +
												  "sheet MUST already exist!\n");
								Console.WriteLine("[Deprecated for dictionary values]\n" +
												  "val 'row' 'collumn' 'number' --> Same as 'add' but the value is overwritten!");
								break;
							}
							case "voice": {
								parser = new DefinitionParser(new System.Text.RegularExpressions.Regex(@"(Mob_)?\w+\.definition"));
								game = new SpeechRecognitionEngine();
								helper = new SpeechRecognitionHelper(ref game);
								enemyHandling = new EnemyHandling();
								helper.OnRecognitionChange += OnRecognitionChange;
								helper.AcquireControl();
								Console.WriteLine("Returned to Main control!");
								mapper.AssignToHotkey(Keys.F1, "voice");
								mapper.AssignToHotkey(Keys.F2, "chest");
								mapper.AssignToHotkey(Keys.F3, "help");
								mapper.AssignToHotkey(Keys.F4, "quit");
								mapper.AssignToHotkey(Keys.F8, KeyModifiers.Shift, "wipe");
								helper.CleanUp();
								enemyHandling.CleanUp();
								enemyHandling = null;
								helper = null;
								game.SpeechRecognized -= Game_ModifierRecognized_Wrapper;
								game.SpeechRecognized -= Game_SpeechRecognized_Wrapper;
								game = null;
								break;
							}
							case "chest": {
								game = new SpeechRecognitionEngine();
								helper = new Chests.ChestVoiceManager(ref game);
								helper.OnRecognitionChange += OnRecognitionChange;
								helper.AcquireControl();
								Console.WriteLine("Returned to Main control!");
								helper = null;
								break;
							}
							case "clear": {
								Console.Clear();
								Console.WriteLine("Welcome to Metin2 siNDiCATE Drop logger");
								Console.WriteLine("Type 'help' for more info on how to use this program");
								break;
							}
							case "wipe": {
								Console.WriteLine("Wiping data...");
								if (File.Exists(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg")) {
									File.Delete(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg");
									Console.WriteLine("config.cfg");
								}
								if (File.Exists(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MobAsociatedDrops.MOB_DROPS_FILE)) {
									File.Delete(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + MobAsociatedDrops.MOB_DROPS_FILE);
									Console.WriteLine(MobAsociatedDrops.MOB_DROPS_FILE);
								}
								if (File.Exists(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + Configuration.FILE_NAME)) {
									File.Delete(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + Configuration.FILE_NAME);
									Console.WriteLine(Configuration.FILE_NAME);
								}
								config = new Configuration(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg");
								interaction = new SpreadsheetInteraction(config.xlsxFile);
								break;
							}
							default: {
								Console.WriteLine("Not a valid command, type 'help' for more info");
								break;
							}
						}
						break;
					}
					case 2: {
						switch (commandBlocks[0]) {
							case "sheet": {
								string sheet = commandBlocks[1];
								interaction.OpenWorksheet(sheet);
								break;
							}
							case "voice": {
								if (commandBlocks[1] == "debug") {
									debug = true;
									Console.WriteLine("Entering debug mode");
									parser = new DefinitionParser(new System.Text.RegularExpressions.Regex(".+"));
									game = new SpeechRecognitionEngine();
									helper = new SpeechRecognitionHelper(ref game);
									enemyHandling = new EnemyHandling();
									helper.OnRecognitionChange += OnRecognitionChange;
									Console.WriteLine(game.Grammars.Count);

									helper.AcquireControl();

									Console.WriteLine("Returned to Main control!");
									enemyHandling = null;
									helper = null;
									game.SpeechRecognized -= Game_ModifierRecognized_Wrapper;
									game.SpeechRecognized -= Game_SpeechRecognized_Wrapper;
								}
								break;
							}
							default: {
								Console.WriteLine("Not a valid command, type 'help' for more info");
								break;
							}
						}
						break;
					}
					case 3: {
						switch (commandBlocks[0]) {
							case "sheet": {
								if (commandBlocks[1] == "add") {
									interaction.OpenWorksheet(commandBlocks[2]);
									Console.WriteLine("Added sheet " + commandBlocks[2]);
								}
								break;
							}
						}
						break;
					}
					case 4: {
						switch (commandBlocks[0]) {
							//Add a number to a cell, args: row, col, number
							case "add": {
								int successCounter = 0;
								if (int.TryParse(commandBlocks[1], out int row)) {
									successCounter += 1;
								}
								if (int.TryParse(commandBlocks[2], out int collum)) {
									successCounter += 1;
								}
								if (int.TryParse(commandBlocks[3], out int numberToAdd)) {
									successCounter += 1;
								}
								if (successCounter == 3) {
									interaction.AddNumberTo(new ExcelCellAddress(collum, row), numberToAdd);
									Console.WriteLine("Added " + numberToAdd + " to sheet at [" + row + "," + collum + "]");
								}
								break;
							}
							case "val": {
								int successCounter = 0;
								if (int.TryParse(commandBlocks[1], out int row)) {
									successCounter += 1;
								}
								if (int.TryParse(commandBlocks[2], out int collum)) {
									successCounter += 1;
								}
								if (successCounter == 2) {
									interaction.InsertValue(new ExcelCellAddress(collum, row), commandBlocks[3]);
									Console.WriteLine("Added " + commandBlocks[3] + " to sheet at [" + row + "," + collum + "]");
								}
								break;
							}
							default: {
								Console.WriteLine("Not a valid command, type 'help' for more info");
								break;
							}
						}
						break;
					}
				}
				currCommand = "";
				mapper.FreeGame();
			}
		}

		private static string HandleHotkeyCommand(string command) {
			Console.WriteLine(command);
			return command;
		}

		private static void OnRecognitionChange(RecognitionState state) {
			switch (state) {
				case RecognitionState.ERROR: {
					Console.WriteLine("Something went wrong");
					break;
				}
				case RecognitionState.CONTROL_RUNNING: {
					game.SpeechRecognized -= Game_SpeechRecognized_Wrapper;
					game.RecognizeAsyncStop();
					break;
				}
				case RecognitionState.DROP_LOGGER_RUNNING: {
					game.SetInputToDefaultAudioDevice();
					game.SpeechRecognized += Game_SpeechRecognized_Wrapper;
					game.RecognizeAsync(RecognizeMode.Multiple);
					break;
				}
			}
		}

		private static void Game_SpeechRecognized_Wrapper(object sender, SpeechRecognizedEventArgs e) {
			Game_SpeechRecognized(sender, new SpeechRecognizedArgs(e.Result.Text, e.Result.Confidence));
		}
		private static void Game_SpeechRecognized(object sender, SpeechRecognizedArgs e) {
			//TODO bind specific items to hotkeys
			if (SpeechRecognitionHelper.reverseModifierDict.ContainsKey(e.text)) {
				SpeechRecognitionHelper.ModifierWords current = SpeechRecognitionHelper.reverseModifierDict[e.text];
				switch (current) {
					case SpeechRecognitionHelper.ModifierWords.NEW_TARGET: {
						game.Grammars[0].Enabled = false;
						game.Grammars[2].Enabled = true;
						OnModifierWordHear?.Invoke(SpeechRecognitionHelper.currentModifier, "");
						Console.WriteLine("Listening for enemy type");
						SpeechRecognitionHelper.currentModifier = SpeechRecognitionHelper.ModifierWords.NEW_TARGET;
						game.SpeechRecognized += Game_ModifierRecognized_Wrapper;
						game.SpeechRecognized -= Game_SpeechRecognized_Wrapper;
						return;
					}
					case SpeechRecognitionHelper.ModifierWords.UNDO: {
						Console.Write("Undoing...");
						Confirmation.SelectivelyDisableEnableGrammars(ref game, true);
						OnModifierWordHear?.Invoke(SpeechRecognitionHelper.ModifierWords.UNDO, "");
						Confirmation.SelectivelyDisableEnableGrammars(ref game, false);
						return;
					}
					case SpeechRecognitionHelper.ModifierWords.TARGET_KILLED: {
						Console.Write("Switching back to default NONE target");
						OnModifierWordHear?.Invoke(SpeechRecognitionHelper.ModifierWords.TARGET_KILLED, "");
						return;
					}
					default: {
						return;
					}
				}
			}
			Console.WriteLine(e.text + " -- " + e.confidence);
			enemyHandling.ItemDropped(e.text);
		}

		private static void Game_ModifierRecognized_Wrapper(object sender, SpeechRecognizedEventArgs e) {
			Game_ModifierRecognized(sender, new SpeechRecognizedArgs(e.Result.Text, e.Result.Confidence));
		}
		private static void Game_ModifierRecognized(object sender, SpeechRecognizedArgs e) {
			switch (SpeechRecognitionHelper.currentModifier) {
				case SpeechRecognitionHelper.ModifierWords.NEW_TARGET: {
					bool recognisedWordComesFromGrammarOne = false;
					foreach (string theString in SpeechRecognitionHelper.modifierDict.Values) {
						if (theString == e.text) {
							recognisedWordComesFromGrammarOne = true;
							Console.WriteLine("NewTarget cancelled");
						}
					}
					if (!recognisedWordComesFromGrammarOne) {
						OnModifierWordHear?.Invoke(SpeechRecognitionHelper.currentModifier, e.text);
					}
					game.Grammars[0].Enabled = true;
					game.Grammars[2].Enabled = false;
					SpeechRecognitionHelper.currentModifier = SpeechRecognitionHelper.ModifierWords.NONE;
					break;
				}
			}
			game.SpeechRecognized -= Game_ModifierRecognized_Wrapper;
			game.SpeechRecognized += Game_SpeechRecognized_Wrapper;
		}
	}
}

