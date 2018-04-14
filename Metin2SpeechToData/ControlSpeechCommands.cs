using System;
using System.IO;

namespace Metin2SpeechToData {
	class ControlSpeechCommands {

		public const string START = "START";
		public const string PAUSE = "PAUSE";
		public const string STOP = "STOP";
		public const string SWITCH = "SWITCH";

		public string getStartCommand { get; private set; }
		public string getPauseCommand { get; private set; }
		public string getStopCommand { get; private set; }
		public string getSwitchGrammarCommand { get; private set; }


		public ControlSpeechCommands(string relativeFilePath) {
			using (StreamReader sr = File.OpenText(Directory.GetCurrentDirectory() + relativeFilePath + ".definition")) {
				short current = 0;
				while (!sr.EndOfStream) {
					string line = sr.ReadLine();
					if (line.StartsWith("#") || string.IsNullOrWhiteSpace(line)) {
						continue;
					}
					switch (current) {
						case 0: {
							getStartCommand = line;
							break;
						}
						case 1: {
							getPauseCommand = line;
							break;
						}
						case 2: {
							getStopCommand = line;
							break;
						}
						case 3: {
							getSwitchGrammarCommand = line;
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
