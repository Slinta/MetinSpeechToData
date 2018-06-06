﻿using System;
using System.IO;
using System.Windows.Forms;
using System.Threading;

namespace Metin2SpeechToData {
	public static class Program {

		public static bool debug { get; private set; } = false;

		private static DefinitionParser parser { get; set; }

		public static SpreadsheetInteraction interaction { get; private set; }
		public static Configuration config { get; private set; }

		public static HotKeyMapper mapper { get; private set; }

		public static string currCommand { get; set; } = "";

		[STAThread]
		static void Main(string[] args) {
			config = new Configuration(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg");
			interaction = new SpreadsheetInteraction(config.xlsxFile);
			Confirmation.Initialize();
			mapper = new HotKeyMapper();
			parser = new DefinitionParser();

			AssignInitialHotkeys();
			bool continueRunning = true;
			//
			Console.WriteLine("Welcome to Metin2 siNDiCATE Drop logger");
			Console.WriteLine("Type 'help' for more info on how to use this program");
			while (continueRunning) {

				Console.WriteLine("Commands:" +
					"\n(F1)-Drop with voice recognition" +
					"\n(F2)-Chests" +
					"\n(F3)-Help" +
					"\n(F4)-Quit" +
					"\n(Shift + F8)-Wipe files(see 'help')");

				string command = Console.ReadLine();

				if (currCommand != "") {
					command = currCommand;
				}

				string[] commandBlocks = command.Split(' ');

				//Switch over length
				switch (commandBlocks.Length) {
					case 1: {
						switch (commandBlocks[0]) {
							case "quit":
							case "exit": {
								if (Confirmation.WrittenConfirmation("Do you want to quit ?")) {
									Environment.Exit(0);
								}
								break;
							}
							case "help": {
								Console.WriteLine("Existing commands:");
								Console.WriteLine("quit / exit --> Close the application\n");
								Console.WriteLine("clear --> Clears the console\n");
								Console.WriteLine("voice --> Enables voice control without/with debug prints\n");
								Console.WriteLine("chest --> Opens speech recognition for chest drops (Gold/Silver[+-])\n");
								Console.WriteLine("wipe --> removes main sheet and all nonessential custom data !CAUTION ADVISED!");
								break;
							}
							case "voice": {
								GameRecognizer gameRecognizer = new GameRecognizer();
								gameRecognizer.helper.AcquireControl();
								Console.WriteLine("Returned to Main control!");
								mapper.FreeGameHotkeys();
								AssignInitialHotkeys();
								break;
							}
							case "chest": {
								ChestRecognizer chestRecognizer = new ChestRecognizer();
								chestRecognizer.helper.AcquireControl();
								Console.WriteLine("Returned to Main control!");
								mapper.FreeGameHotkeys();
								AssignInitialHotkeys();
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
								Console.Write("config.cfg");

								if (File.Exists(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg")) {
									File.Delete(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg");
									Console.WriteLine("...SUCCESS");
								}
								else {
									Console.WriteLine("...file already deleted");
								}

								Console.Write(Configuration.DEFAULT_FILE_NAME);
								if (File.Exists(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + Configuration.DEFAULT_FILE_NAME)) {
									File.Delete(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + Configuration.DEFAULT_FILE_NAME);
									Console.WriteLine("...SUCCESS");
								}
								else {
									Console.WriteLine("...file already deleted");
								}
								Console.WriteLine();
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
				}
				currCommand = "";
				mapper.FreeGameHotkeys();
			}
		}

		private static void AssignInitialHotkeys() {
			mapper.AssignToHotkey(Keys.F1, "voice");
			mapper.AssignToHotkey(Keys.F2, "chest");
			mapper.AssignToHotkey(Keys.F3, "help");
			mapper.AssignToHotkey(Keys.F4, "quit");
			mapper.AssignToHotkey(Keys.F8, KeyModifiers.Shift, "wipe");
		}

		//private static string GetCommand(ConsoleKeyInfo info) {
		//	switch (info.Key) {
		//		case ConsoleKey.F1: {
		//			return "voice";
		//		}
		//		case ConsoleKey.F2: {
		//			return "chests";
		//		}
		//		case ConsoleKey.F3: {
		//			return "help";
		//		}
		//		case ConsoleKey.F4: {
		//			return "quit";
		//		}
		//		case ConsoleKey.F8: {
		//			return "wipe";
		//		}
		//	}
		//	throw new InvalidDataException();
		//}

		//private static void ClearLastLine(string text) {
		//	Console.SetCursorPosition(text.Length, Console.CursorTop);
		//	Console.Write(new string(' ', Console.BufferWidth));
		//	Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop - 1);
		//}
	}
}

