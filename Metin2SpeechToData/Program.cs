using System;
using System.IO;
using System.Speech.Recognition;
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

		private static SpeechRecognitionEngine game;
		private static SpeechRecognitionHelper helper;

		private static DefinitionParser parser;
		private static EnemyHandling enemyHandling;

		public static SpreadsheetInteraction interaction;


		public static bool debug = false;
		private static WrittenControl debugControl;

		[STAThread]
		static void Main(string[] args) {
			Console.WriteLine("Welcome to Metin2 siNDiCATE Drop logger");
			Console.WriteLine("Type 'help' for more info on how to use this program");
			bool continueRunning = true;
			interaction = new SpreadsheetInteraction(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Data.xlsx", "Sheet1");
			while (continueRunning) {
				string command = Console.ReadLine();

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
								Console.WriteLine("Existing commands:");
								Console.WriteLine("quit / exit --> Close the application");
								Console.WriteLine("voice / voice debug --> Enables voice control without/with debug prints");
								Console.WriteLine("file 'location' --> Creates a new file at 'location' <- replace with valid path");
								Console.WriteLine("file 'location' 'sheet name' --> Creates a new file at 'location' <- replace with valid path, with a sheet 'sheet name' inside");
								Console.WriteLine("sheet 'sheet name' --> Swithches current working sheet to 'sheet name' <- replace with valid sheet name, sheet MUST already exist(by typing \"file 'location' 'sheet name'\")");
								Console.WriteLine("add 'row' 'collumn' 'number' --> Adds value of 'number' to currently open sheet at 'row' 'collumn', sheet MUST already exist");
								Console.WriteLine("val 'row' 'collumn' 'number' --> Same as 'add' but the value is overwritten!");
								break;
							}
							case "voice": {
								parser = new DefinitionParser();
								game = new SpeechRecognitionEngine();
								helper = new SpeechRecognitionHelper(ref game);
								enemyHandling = new EnemyHandling();
								helper.OnRecognitionChange += OnRecognitionChange;
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
							case "file": {
								string location = commandBlocks[1];
								if (commandBlocks[1] == "default") {
									location = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Data.xlsx";
								}
								interaction = new SpreadsheetInteraction(location);
								break;
							}
							//Switch current working sheet, sheet must already exist!
							case "sheet": {
								string sheet = commandBlocks[1];
								interaction.OpenWorksheet(sheet);
								break;
							}
							case "voice": {
								if (commandBlocks[1] == "debug") {
									debug = true;
								}
								parser = new DefinitionParser();
								game = new SpeechRecognitionEngine();
								Console.WriteLine(game.Grammars.Count);
								helper = new SpeechRecognitionHelper(ref game);
								helper.OnRecognitionChange += OnRecognitionChange;
								enemyHandling = new EnemyHandling();
								debugControl = new WrittenControl(helper.controlCommands, ref enemyHandling);
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
							case "file": {
								string location = commandBlocks[1];
								string sheet = commandBlocks[2];
								if (commandBlocks[1] == "default") {
									location = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Data.xlsx";
								}
								if (commandBlocks[2] == "default") {
									sheet = "Data";
								}
								interaction = new SpreadsheetInteraction(location, sheet);
								break;
							}
							default: {
								Console.WriteLine("Not a valid command, type 'help' for more info");
								break;
							}
						}
						break;
					}
					case 4: {
						switch (commandBlocks[0]) {
							//Add a number to a cell, args: row, col, number
							case "add": {
								if (interaction == null) {
									throw new Exception("File does not exist");
								}
								int successCounter = 0;
								if (int.TryParse(commandBlocks[1], out int row)) {
									successCounter += 1;
								}
								if (int.TryParse(commandBlocks[2], out int collum)) {
									successCounter += 1;
								}
								if (int.TryParse(commandBlocks[3], out int add)) {
									successCounter += 1;
								}
								if (successCounter == 3) {
									interaction.AddNumberTo(new ExcelCellAddress(collum, row), add);
								}
								break;
							}
							case "val": {
								if (interaction == null) {
									throw new Exception("File does not exist");
								}
								int successCounter = 0;
								if (int.TryParse(commandBlocks[1], out int row)) {
									successCounter += 1;
								}
								if (int.TryParse(commandBlocks[2], out int collum)) {
									successCounter += 1;
								}
								if (successCounter == 2) {
									interaction.InsertValue(new ExcelCellAddress(collum, row), commandBlocks[3]);
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
			}
		}

		private static void OnRecognitionChange(RecognitionState state) {
			switch (state) {
				case RecognitionState.ERROR: {
					Console.WriteLine("Something went wrong");
					break;
				}
				case RecognitionState.CONTROL_RUNNING: {
					game.SpeechRecognized -= Game_SpeechRecognized;
					game.RecognizeAsyncStop();
					break;
				}
				case RecognitionState.DROP_LOGGER_RUNNING: {
					game.SetInputToDefaultAudioDevice();
					game.SpeechRecognized += Game_SpeechRecognized;
					game.RecognizeAsync(RecognizeMode.Multiple);
					break;
				}
				default: {
					break;
				}
			}
		}

		private static void Game_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			foreach (string s in SpeechRecognitionHelper.modifierDict.Values) {
				if (s == e.Result.Text) {
					switch (s) {
						case "New Target": {
							game.Grammars[0].Enabled = false;
							game.Grammars[2].Enabled = true;
							OnModifierWordHear?.Invoke(SpeechRecognitionHelper.currentModifier, "");
							Console.WriteLine("Listening for enemy type");
							SpeechRecognitionHelper.currentModifier = SpeechRecognitionHelper.ModifierWords.NEW_TARGET;
							break;
						}
						case "Undo": {
							Console.Write("Undoing...");
							OnModifierWordHear?.Invoke(SpeechRecognitionHelper.ModifierWords.UNDO, "");
							return;
						}
					}
				}
			}
			if (SpeechRecognitionHelper.currentModifier != SpeechRecognitionHelper.ModifierWords.NONE) {
				game.SpeechRecognized += Game_ModifierRecognized;
				game.SpeechRecognized -= Game_SpeechRecognized;
				return;
			}
			enemyHandling.ItemDropped(e.Result.Text);
			Console.WriteLine(e.Result.Text + " -- " + e.Result.Confidence);
		}

		private static void Game_ModifierRecognized(object sender, SpeechRecognizedEventArgs e) {
			switch (SpeechRecognitionHelper.currentModifier) {
				case SpeechRecognitionHelper.ModifierWords.NEW_TARGET: {
					OnModifierWordHear?.Invoke(SpeechRecognitionHelper.currentModifier, e.Result.Text);
					game.Grammars[0].Enabled = true;
					game.Grammars[2].Enabled = false;
					SpeechRecognitionHelper.currentModifier = SpeechRecognitionHelper.ModifierWords.NONE;
					break;
				}
			}
			game.SpeechRecognized -= Game_ModifierRecognized;
			game.SpeechRecognized += Game_SpeechRecognized;
		}
	}
}

