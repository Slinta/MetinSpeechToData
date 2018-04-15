using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Recognition;
using System.IO;
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

		public static bool debug = false;

		static void Main(string[] args) {
			//Welcome
			Console.WriteLine("Welcome");
			Console.WriteLine("This is a project");

			//exit condition
			bool continueRunning = true;

			//the object containing the current spreadsheet, excel file and the methods to alter it
			SpreadsheetInteraction interaction = null;
			//main program
			while (continueRunning) {
				//check for commands
				string command = Console.ReadLine();
				//divide the string onto blocks separated by spacebar
				string[] commandBlocks = command.Split((char)ConsoleKey.Spacebar);
				switch (commandBlocks.Length) {
					//One word command
					case 1: {
						switch (commandBlocks[0]) {
							//ask for confirmation and then exit the application by stoping while
							case "quit": {
								Console.WriteLine("Do you want to quit? y/n");
								if (Console.ReadKey().Key == ConsoleKey.Y) {
									continueRunning = false;
								}
								else {
									Console.WriteLine("Exit aborted");
								}
								break;
							}
							//list commands
							case "help": {
								Console.WriteLine("commands are:");
								Console.WriteLine("add collom row number");
								Console.WriteLine("quit");
								break;
							}
							default: {


								break;
							}
						}
						break;
					}
					case 2: {
						//switch by  firs word in command
						switch (commandBlocks[0]) {
							//change the edited file
							case "file": {
								string location = commandBlocks[1];
								//if the location is default use this
								if (commandBlocks[1] == "default") {
									location = "\\\\SLINTA-PC\\Sharing\\Metin2\\BokjungData.xlsx";
								}
								interaction = new SpreadsheetInteraction(location);
								break;
							}
							//change the sheet in the edited file, sheet must already exist
							case "sheet": {
								if (interaction == null) {
									Console.WriteLine("You have to assign a file before you assign the sheet");
									break;
								}
								string sheet = commandBlocks[1];
								interaction.OpenWorksheet(sheet);
								break;
							}

							case "vr": {
								if (commandBlocks[1] == "debug") {
									debug = true;
								}

								parser = new DefinitionParser();
								game = new SpeechRecognitionEngine();
								if (debug) {
									Console.WriteLine(game.Grammars.Count);
								}
								helper = new SpeechRecognitionHelper(ref game);
								helper.OnRecognitionChange += OnRecognitionChange;
								break;
							}
						}
						break;
					}
					case 3: {
						switch (commandBlocks[0]) {
							//specify both file and sheet
							case "file": {
								string location = commandBlocks[1];
								string sheet = commandBlocks[2];
								if (commandBlocks[1] == "default") {
									location = "\\\\SLINTA-PC\\Sharing\\Metin2\\BokjungData.xlsx";
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
									Console.WriteLine("No file set yet");
								}
								int row;
								int collum;
								int add;
								int successCounter = 0;
								if (int.TryParse(commandBlocks[1], out row)) {
									successCounter += 1;
								}
								if (int.TryParse(commandBlocks[2], out collum)) {
									successCounter += 1;
								}
								if (int.TryParse(commandBlocks[3], out add)) {
									successCounter += 1;
								}
								if (successCounter == 3) {

									interaction.AddNumberTo(new ExcelCellAddress(collum, row), add);
								}
								break;
							}
							default: {
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
		}
	}
}

