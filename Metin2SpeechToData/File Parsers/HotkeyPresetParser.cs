using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Metin2SpeechToData.Structures;

namespace Metin2SpeechToData {
	public class HotkeyPresetParser {

		private readonly FileInfo[] hotkeyFiles;

		public List<int> activeKeyIDs { get; private set; }

		public List<Keys> currentCustomKeys { get; private set; }
		private List<(Keys key, string name)> _hotkeySelection;

		public HotkeyPresetParser(string selectedArea) {
			DirectoryInfo directory = new DirectoryInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Hotkeys");
			hotkeyFiles = directory.GetFiles("* *.definition");
			activeKeyIDs = new List<int>();
			currentCustomKeys = new List<Keys>();
			Load(selectedArea);
		}


		public void SetKeysActiveState(bool state) {
			foreach (Keys key in currentCustomKeys) {
				Program.mapper.SetInactive(key, !state);
			}
		}

		public void Selected_Hotkey(SpeechRecognizedArgs args) {
			Console.WriteLine("Selected " + args.text);
			for (int i = 0; i < _hotkeySelection.Count; i++) {
				if (_hotkeySelection[i].name == args.text) {
					Program.mapper.SetInactive(_hotkeySelection[i].key, true);
				}
			}
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
						currentCustomKeys.Add(key);
						break;
					}
					case 2: {
						bool one = Enum.TryParse(curr[0], out KeyModifiers mod1);
						Keys key = (Keys)Enum.Parse(typeof(Keys), curr[1]);
						if (one) {
							int ID = Program.mapper.AssignToHotkey(key, mod1, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
							Console.WriteLine("Assigning - " + mod1 + " + " + key + " for item " + itemName);
							activeKeyIDs.Add(ID);
							currentCustomKeys.Add(key);
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
							currentCustomKeys.Add(key);
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
						if (one && two && three) {
							int ID = Program.mapper.AssignToHotkey(key, mod1, mod2, mod3, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
							Console.WriteLine("Assigning - " + mod1 + " + " + mod2 + " + " + mod3 + " + " + key + " for item " + itemName);
							activeKeyIDs.Add(ID);
							currentCustomKeys.Add(key);
						}
						else {
							Console.WriteLine("Unable to parse one of '" + curr[0] + " | " + curr[1] + " | " + curr[2] + "'");
						}
						break;
					}
				}
			}
			SetKeysActiveState(false);
		}

		public void Load(string area) {
			FileInfo[] validForArea = hotkeyFiles.Where(file => new Regex(Regex.Escape(area) + @"\ \d+\.definition").IsMatch(file.Name)).ToArray();
			_hotkeySelection = new List<(Keys key, string name)>();
			if (validForArea.Length > 0) {
				Console.WriteLine("Found some hotkey definitions for " + area + ", load them ?");
				for (int i = 0; i < validForArea.Length; i++) {
					Console.WriteLine("(" + (i + 1) + ") " + validForArea[i].Name.Split('.')[0]);
					Program.mapper.AssignToHotkey((Keys)((int)Keys.D1 + i), Selected_Hotkey, new SpeechRecognizedArgs(validForArea[i].Name.Split(':')[0], 100));
					_hotkeySelection.Add(((Keys)((int)Keys.D1 + i), validForArea[i].Name.Split(':')[0]));
				}
			}
			else {
				Console.WriteLine("No hotkey mappings were found, continuing.");
			}
		}
	}
}
