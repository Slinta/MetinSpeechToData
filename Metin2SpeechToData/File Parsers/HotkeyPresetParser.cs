using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Metin2SpeechToData.Structures;
using System.Threading;

namespace Metin2SpeechToData {
	public class HotkeyPresetParser : IDisposable {

		private readonly FileInfo[] hotkeyFiles;

		private readonly List<int> _areas = new List<int>();

		private readonly ManualResetEventSlim evnt = new ManualResetEventSlim();


		public HotkeyPresetParser(string selectedArea) {
			DirectoryInfo directory = new DirectoryInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Hotkeys");
			hotkeyFiles = directory.GetFiles("* *.definition");

			Thread t = new Thread(delegate () {
				Load(selectedArea);
				evnt.Wait();
				foreach (int item in _areas) {
					Program.mapper.FreeSpecific(item);
				}
				_areas.Clear();
				evnt.Dispose();
			});
			t.IsBackground = true;
			t.Name = "My Thread";
			t.Start();
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
						Program.mapper.AssignItemHotkey(key, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
						Console.WriteLine("Assigning - " + key + " for item " + itemName);
						break;
					}
					case 2: {
						bool one = Enum.TryParse(curr[0], out KeyModifiers mod1);
						Keys key = (Keys)Enum.Parse(typeof(Keys), curr[1]);
						Program.mapper.AssignItemHotkey(key, mod1, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
						if (one) {
							Console.WriteLine("Assigning - " + mod1 + " + " + key + " for item " + itemName);
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
							Program.mapper.AssignItemHotkey(key, mod1, mod2, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
							Console.WriteLine("Assigning - " + mod1 + " + " + mod2 + " + " + key + " for item " + itemName);
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
							Program.mapper.AssignItemHotkey(key, mod1, mod2, mod3, Program.mapper.EnemyHandlingItemDroppedWrapper, new SpeechRecognizedArgs(itemName, 100));
							Console.WriteLine("Assigning - " + mod1 + " + " + mod2 + " + " + mod3 + " + " + key + " for item " + itemName);
						}
						else {
							Console.WriteLine("Unable to parse one of '" + curr[0] + " | " + curr[1] + " | " + curr[2] + "'");
						}
						break;
					}
				}
			}
			evnt.Set();
		}

		public void Load(string area) {
			FileInfo[] validForArea = hotkeyFiles.Where(file => new Regex(Regex.Escape(area) + @"\ \d+\.definition").IsMatch(file.Name)).ToArray();
			if (validForArea.Length > 0) {
				Console.WriteLine("\nFound some hotkey definitions for " + area + ", load them ?");
				for (int i = 0; i < validForArea.Length; i++) {
					Console.WriteLine("(" + (i + 1) + ") " + validForArea[i].Name.Split('.')[0]);
					int unregID = Program.mapper.AssignToHotkey((Keys)((int)Keys.D1 + i), Selected_Hotkey, new SpeechRecognizedArgs(validForArea[i].Name.Split(':')[0], 100));
					_areas.Add(unregID);
				}
			}
			else {
				Console.WriteLine("No hotkey mappings were found, continuing.");
			}
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				evnt.Dispose();
				disposedValue = true;
			}
		}

		~HotkeyPresetParser() {
			Dispose(false);
		}

		// This code added to correctly implement the disposable pattern.
		public void Dispose() {
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
