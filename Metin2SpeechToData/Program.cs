﻿using System;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public static class Program {

		public static bool debug { get; private set; } = false;

		public static GameRecognizer gameRecognizer { get; private set; }
		private static DefinitionParser parser { get; set; }

		public static SpreadsheetInteraction interaction { get; private set; }
		public static Configuration config { get; private set; }
		public static ControlSpeechCommands controlCommands { get; private set; }

		public static HotKeyMapper mapper { get; private set; }

		public static string currCommand { get; set; } = "";

		/* NIY List:
		 * TODO: Sessions
		 * TODO: Error check and speech recognition overlaps
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
							case "quit":
							case "exit": {
								Console.WriteLine("Do you want to quit? y/n");
								if (Console.ReadKey().Key == ConsoleKey.Y) {
									continueRunning = false;
								}
								break;
							}
							case "help": {
								Console.WriteLine("Existing commands:");
								Console.WriteLine("quit / exit --> Close the application\n");
								Console.WriteLine("clear --> Clears the console\n");
								Console.WriteLine("voice / voice debug --> Enables voice control without/with debug prints\n");
								Console.WriteLine("chest --> Opens speech recognition for chest drops (Gold/Silver[+-])\n");
								Console.WriteLine("wipe --> removes the sheet and all custom data !CAUTION ADVISED!");
								break;
							}
							case "voice": {
								parser = new DefinitionParser(new System.Text.RegularExpressions.Regex(@"(Mob_)?\w+\.definition"));
								gameRecognizer = new GameRecognizer();
								gameRecognizer.helper.AcquireControl();
								Console.WriteLine("Returned to Main control!");
								mapper.FreeGameHotkeys();
								mapper.AssignToHotkey(Keys.F1, "voice");
								mapper.AssignToHotkey(Keys.F2, "chest");
								mapper.AssignToHotkey(Keys.F3, "help");
								mapper.AssignToHotkey(Keys.F4, "quit");
								mapper.AssignToHotkey(Keys.F8, KeyModifiers.Shift, "wipe");
								break;
							}
							case "chest": {
								parser = new DefinitionParser(new System.Text.RegularExpressions.Regex(@"\w+\ (C|c)hest[+-]?\.definition"));
								ChestRecognizer chestRecognizer = new ChestRecognizer();
								chestRecognizer.helper.AcquireControl();
								Console.WriteLine("Returned to Main control!");
								mapper.FreeGameHotkeys();
								mapper.AssignToHotkey(Keys.F1, "voice");
								mapper.AssignToHotkey(Keys.F2, "chest");
								mapper.AssignToHotkey(Keys.F3, "help");
								mapper.AssignToHotkey(Keys.F4, "quit");
								mapper.AssignToHotkey(Keys.F8, KeyModifiers.Shift, "wipe");
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
							case "voice": {
								if (commandBlocks[1] == "debug") {
									debug = true;
									Console.WriteLine("Entering debug mode");
									parser = new DefinitionParser(new System.Text.RegularExpressions.Regex(".+"));
									gameRecognizer = new GameRecognizer();
									gameRecognizer.helper.AcquireControl();
									Console.WriteLine("Returned to Main control!");
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
				mapper.FreeGameHotkeys();
			}
		}
	}
}

