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

		private static SpeechRecognitionEngine game;
		private static DefinitionParser parser;
		private static SpeechRecognitionHelper helper;
		//The object containing current spreadsheet, excel file and the methods to alter it
		public static SpreadsheetInteraction interaction;
		public static bool debug = false;
		
		static void Main(string[] args) {
			Console.WriteLine("Welcome to Metin2 siNDiCATE Drop logger");

			bool continueRunning = true;
	
			

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
			Console.WriteLine(e.Result.Text + " -- " + e.Result.Confidence);
			Program.interaction.MakeANewSpreadsheet(DefinitionParser.instance.currentGrammarFile);
		}
	}
}

