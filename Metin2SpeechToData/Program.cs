using System;
using System.Speech.Recognition;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	class Program {
		public enum RecognitionState {
			ERROR,
			CONTROL_RUNNING,
			DROP_LOGGER_RUNNING,
			SWITCHING,
		}

		public delegate void Recognition(RecognitionState state);
		public delegate void ModifierTrigger(string word, params string[] args);
		public static event ModifierTrigger OnModifierWordHear;

		private static EnemyHandling enemyHandling;
		private static SpeechRecognitionEngine game;
		private static DefinitionParser parser;
		private static SpeechRecognitionHelper helper;
		//The object containing current spreadsheet, excel file and the methods to alter it
		public static SpreadsheetInteraction interaction;
		public static bool debug = false;

		public static readonly string[] modifierWords = new string[1] { "New Target" };
		private static bool listeningForArgument = false;
		
		static void Main(string[] args) {
			Console.WriteLine("Welcome to Metin2 siNDiCATE Drop logger");
			enemyHandling = new EnemyHandling();
			bool continueRunning = true;
			interaction = new SpreadsheetInteraction(@"\\SLINTA-PC\Sharing\Metin2\BokjungData.xlsx", "Data");
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
								break;
							}
							case "voice": {
								parser = new DefinitionParser();
								game = new SpeechRecognitionEngine();
								helper = new SpeechRecognitionHelper(ref game);
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
									location = @"\\SLINTA-PC\Sharing\Metin2\BokjungData.xlsx";
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
									location = @"\\SLINTA-PC\Sharing\Metin2\BokjungData.xlsx";
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
							//add a number to a cell, args: collon, row, number
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
									interaction.InsertText(new ExcelCellAddress(collum, row), commandBlocks[3]);
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
					game.SpeechRecognized += Game_ModifierArgument;
					game.RecognizeAsync(RecognizeMode.Multiple);
					break;
				}
				default: {
					break;
				}
			}
		}

		private static string currentModifier;

		private static void Game_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			foreach (string s in modifierWords) {
				if(s == e.Result.Text) {
					switch (s) {
						case "New Target": {
							game.Grammars[0].Enabled = false;
							game.Grammars[2].Enabled = true;
							listeningForArgument = true;
							Console.WriteLine("Listening for enemy type");
							currentModifier = s;
							Console.WriteLine(e.Result.Text + " -- " + e.Result.Confidence);
							return;
						}
						
					}
				}
			}
			Console.WriteLine(e.Result.Text + " -- " + e.Result.Confidence);
		}

		private static void Game_ModifierArgument(object sender, SpeechRecognizedEventArgs e) {
			if (listeningForArgument && e.Result.Text != "New Target") {
				enemyHandling.EnemyFinished();
				listeningForArgument = false;
				Console.WriteLine(e.Result.Text);
				OnModifierWordHear?.Invoke(currentModifier, e.Result.Text);
				game.Grammars[0].Enabled = true;
				game.Grammars[2].Enabled = false;
			}
			else if(!listeningForArgument){
				enemyHandling.Drop(e.Result.Text);
			}
		}
	}
}

