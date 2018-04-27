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
				while (!sr.EndOfStream) {
					string line = sr.ReadLine();
					if (line.StartsWith("#") || string.IsNullOrWhiteSpace(line)) {
						continue;
					}
					string[] split = line.Split(':');
					string modified = split[1].Remove(0, 1);
					switch (split[0]) {
						case "START": {
							getStartCommand = modified;
							break;
						}
						case "PAUSE": {
							getPauseCommand = modified;
							break;
						}
						case "STOP": {
							getStopCommand = modified;
							break;
						}
						case "SWITCH": {
							getSwitchGrammarCommand = modified;
							break;
						}
						default: {
							throw new CustomException("Corrupted Control.definition file, redownload the application.");
						}
					}
				}
			}
		}
	}
}
