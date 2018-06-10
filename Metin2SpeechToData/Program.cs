using System;
using System.IO;
using System.Windows.Forms;

namespace Metin2SpeechToData {
	public static class Program {

		public static bool debug { get; private set; } = false;

		private static DefinitionParser parser { get; set; }

		public static SpreadsheetInteraction interaction { get; private set; }
		public static Configuration config { get; private set; }

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

			new Undo();
			config = new Configuration(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.cfg");
			interaction = new SpreadsheetInteraction(config.xlsxFile);
			Confirmation.Initialize();
			mapper = new HotKeyMapper();
			parser = new DefinitionParser();

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
								Console.WriteLine("Do you want to quit? y/n");
								if (Console.ReadKey(true).Key == ConsoleKey.Y) {
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
								GameRecognizer gameRecognizer = new GameRecognizer();
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
						
								if (File.Exists(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + Configuration.DEFAULT_FILE_NAME)) {
									File.Delete(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + Configuration.DEFAULT_FILE_NAME);
									Console.WriteLine(Configuration.DEFAULT_FILE_NAME);
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
				}
				currCommand = "";
				mapper.FreeGameHotkeys();
			}
		}
	}
}

