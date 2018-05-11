using System.IO;

namespace Metin2SpeechToData {
	public class ControlSpeechCommands {

		//All parsed commands that are available to the controling speech recognizer
		public string getStartCommand { get; private set; }
		public string getPauseCommand { get; private set; }
		public string getStopCommand { get; private set; }
		public string getSwitchGrammarCommand { get; private set; }

		public string getConfirmationCommand { get; private set; }
		public string getRefusalCommand { get; private set; }

		public string getNewTargetCommand { get; private set; }
		public string getTargetKilledCommand { get; private set; }
		public string getUndoCommand { get; private set; }

		public string getHotkeyAssignCommand { get; private set; }
		public string getRemoveTargetCommand { get; private set; }

		public ControlSpeechCommands(string relativeFilePath) {
			if (!File.Exists(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + relativeFilePath)) {
				throw new CustomException("Could not locate 'Control.definition' file! You have to redownload this application");
			}

			using (StreamReader sr = File.OpenText(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + relativeFilePath)) {
				while (!sr.EndOfStream) {
					string line = sr.ReadLine();
					if ( string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) {
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
						case "CONFIRM": {
							getConfirmationCommand = modified;
							break;
						}
						case "REFUSE": {
							getRefusalCommand = modified;
							break;
						}
						case "NEW_TARGET": {
							getNewTargetCommand = modified;
							break;
						}
						case "TARGET_KILLED": {
							getTargetKilledCommand = modified;
							break;
						}
						case "REMOVE_TARGET": {
							getRemoveTargetCommand = modified;
							break;
						}
						case "UNDO": {
							getUndoCommand = modified;
							break;
						}
						case "HOTKEY_ASSIGN": {
							getHotkeyAssignCommand = modified;
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
