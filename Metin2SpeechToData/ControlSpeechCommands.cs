using System;
using System.IO;

namespace Metin2SpeechToData {
	public class ControlSpeechCommands {

		//All parsed commands that are available to the controling speech recognizer
		public string getStartCommand { get; private set; }
		public string getPauseCommand { get; private set; }
		public string getStopCommand { get; private set; }
		public string getSwitchGrammarCommand { get; private set; }

		public ControlSpeechCommands(string relativeFilePath) {
			using (StreamReader sr = File.OpenText(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + relativeFilePath)) {
				short current = 0;
				while (!sr.EndOfStream) {
					string line = sr.ReadLine();
					if (line.StartsWith("#") || string.IsNullOrWhiteSpace(line)) {
						continue;
					}
					string modified = line.Split(':')[1].Remove(0, 1);
					switch (current) {
						case 0: {
							getStartCommand = modified;
							break;
						}
						case 1: {
							getPauseCommand = modified;
							break;
						}
						case 2: {
							getStopCommand = modified;
							break;
						}
						case 3: {
							getSwitchGrammarCommand = modified;
							break;
						}
					}
					current++;
					if (current >= 4) {
						break;
					}
				}
			}
		}
	}
}
