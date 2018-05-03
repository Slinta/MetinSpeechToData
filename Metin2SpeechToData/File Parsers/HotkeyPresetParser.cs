using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System;
using System.Windows.Forms;

namespace Metin2SpeechToData {
	class HotkeyPresetParser {

		private FileInfo[] hotkeyFiles;

		public HotkeyPresetParser(string selectedArea) {
			DirectoryInfo directory = new DirectoryInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Hotkeys");
			hotkeyFiles = directory.GetFiles("* *.definition");
			FileInfo[] validForArea = hotkeyFiles.Where(file => new Regex(selectedArea + @"\ \d+\.definition").IsMatch(file.Name)).ToArray();
			if (validForArea.Length > 0) {
				Console.WriteLine("Found some hotkey definitions for " + selectedArea + ", load them ?");
				for (int i = 0; i < validForArea.Length; i++) {
					Console.WriteLine("(" + (i + 1) + ") " + validForArea[i].Name.Split('.')[0]);
					Program.mapper.AssignToHotkey(Keys.D1, Selected_Hotkey, new SpeechRecognizedArgs(validForArea[i].Name.Split(':')[0], 100));
				}
			}
			else {
				Console.WriteLine("No hotkey mappings were found, continuing.");
			}
		}

		public void Selected_Hotkey(SpeechRecognizedArgs args) {
			Console.WriteLine("Selected " + args.text);
			string[] fileContent = File.ReadAllLines(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Hotkeys" + Path.DirectorySeparatorChar + args.text);
			for (int i = 0; i < fileContent.Length; i++) {
				if (fileContent[i].StartsWith("#") || string.IsNullOrWhiteSpace(fileContent[i])) {
					continue;
				}
				string[] curr = fileContent[i].Split(':');
				string itemName = curr[1].Trim();
				curr = curr[0].Split('+');

				switch (curr.Length) {
					case 1: {
						Keys key = (Keys)Enum.Parse(typeof(Keys), curr[0]);
						Program.mapper.AssignToHotkey(key, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
						break;
					}
					case 2: {
						bool one = Enum.TryParse(curr[0], out KeyModifiers mod1);
						Keys key = (Keys)Enum.Parse(typeof(Keys), curr[1]);
						Program.mapper.AssignToHotkey(key, mod1, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
						break;
					}
					case 3: {
						bool one = Enum.TryParse(curr[0], out KeyModifiers mod1);
						bool two = Enum.TryParse(curr[1], out KeyModifiers mod2);
						Keys key = (Keys)Enum.Parse(typeof(Keys), curr[2]);
						Program.mapper.AssignToHotkey(key, mod1,mod2, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
						break;
					}
					case 4: {
						bool one = Enum.TryParse(curr[0], out KeyModifiers mod1);
						bool two = Enum.TryParse(curr[1], out KeyModifiers mod2);
						bool three = Enum.TryParse(curr[2], out KeyModifiers mod3);
						Keys key = (Keys)Enum.Parse(typeof(Keys), curr[3]);
						//TODO add AssignToHotkey for 3 modifiers
						break;
					}
				}
			}
		}
	}
}
