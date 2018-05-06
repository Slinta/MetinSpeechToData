using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System;
using System.Windows.Forms;
using System.Collections.Generic;

namespace Metin2SpeechToData {
	public class HotkeyPresetParser {

		private FileInfo[] hotkeyFiles;

		public List<int> activeKeyIDs = new List<int>();

		private List<Keys> currKeys = new List<Keys>();

		public HotkeyPresetParser(string selectedArea) {
			DirectoryInfo directory = new DirectoryInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Hotkeys");
			hotkeyFiles = directory.GetFiles("* *.definition");
			Load(selectedArea);
		}


		public void SetKeysActiveState(bool state) {
			foreach (Keys key in currKeys) {
				Program.mapper.SetInactive(key, !state);
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
						int ID = Program.mapper.AssignToHotkey(key, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
						Console.WriteLine("Assigning - " + key + " for item " + itemName);
						activeKeyIDs.Add(ID);
						currKeys.Add(key);
						break;
					}
					case 2: {
						bool one = Enum.TryParse(curr[0], out KeyModifiers mod1);
						Keys key = (Keys)Enum.Parse(typeof(Keys), curr[1]);
						if (one) {
							int ID = Program.mapper.AssignToHotkey(key, mod1, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
							Console.WriteLine("Assigning - " + mod1 + " + " + key + " for item " + itemName);
							activeKeyIDs.Add(ID);
							currKeys.Add(key);
						}
						else {
							Console.WriteLine("Unable to parse '" + curr[0] + "'");
						}
						break;
					}
					case 3: {
						bool one = Enum.TryParse(curr[0], out KeyModifiers mod1);
						bool two = Enum.TryParse(curr[1], out KeyModifiers mod2);
						Keys key = (Keys)Enum.Parse(typeof(Keys), curr[2]);
						if (one && two) {
							int ID = Program.mapper.AssignToHotkey(key, mod1, mod2, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
							Console.WriteLine("Assigning - " + mod1 + " + " + mod2 + " + " + key + " for item " + itemName);
							activeKeyIDs.Add(ID);
							currKeys.Add(key);
						}
						else {
							Console.WriteLine("Unable to parse '" + (one ? curr[1] : curr[0]) + "'");
						}
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
			SetKeysActiveState(false);
		}

		public void Load(string area) {
			FileInfo[] validForArea = hotkeyFiles.Where(file => new Regex(area + @"\ \d+\.definition").IsMatch(file.Name)).ToArray();
			if (validForArea.Length > 0) {
				Console.WriteLine("Found some hotkey definitions for " + area + ", load them ?");
				for (int i = 0; i < validForArea.Length; i++) {
					Console.WriteLine("(" + (i + 1) + ") " + validForArea[i].Name.Split('.')[0]);
					Program.mapper.AssignToHotkey(Keys.D1, Selected_Hotkey, new SpeechRecognizedArgs(validForArea[i].Name.Split(':')[0], 100));
				}
			}
			else {
				Console.WriteLine("No hotkey mappings were found, continuing.");
			}
		}
	}
}
