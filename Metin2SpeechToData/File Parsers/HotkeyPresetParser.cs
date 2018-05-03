using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System;

namespace Metin2SpeechToData {
	class HotkeyPresetParser {

		private FileInfo[] hotkeyFiles;

		private string areaHotkeyFile;

		public HotkeyPresetParser(string selectedArea) {
			DirectoryInfo directory = new DirectoryInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "Hotkeys");
			hotkeyFiles = directory.GetFiles("* *.definition");
			FileInfo[] validForArea = hotkeyFiles.Where(file => new Regex(selectedArea + @"\ \d+\.definition").IsMatch(file.Name)).ToArray();
			if (validForArea.Length > 0) {
				Console.WriteLine("Found some hotkey definitions, load them ?");
				for (int i = 0; i < validForArea.Length; i++) {
					Console.WriteLine("(F" + (i + 1) + ") " + validForArea[i].Name.Split('.')[0]);
				}
			}
			else {
				Console.WriteLine("No hotkey mappings were found, continuing.");
			}
		}
	}
}
