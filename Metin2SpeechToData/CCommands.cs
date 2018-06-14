using System;
using System.IO;

namespace Metin2SpeechToData {
	public class CCommands {

		//All parsed commands that are available to the controling speech recognizer
		public const string CONTROL_DEFINITION_NAME = "Control.definition";

		public static string getStartCommand { get; private set; }
		public static string getPauseCommand { get; private set; }

		public static string getStopCommand { get; private set; }
		public static string getSwitchGrammarCommand { get; private set; }

		public static string getConfirmationCommand { get; private set; }
		public static string getRefusalCommand { get; private set; }

		public static string getNewTargetCommand { get; private set; }
		public static string getTargetKilledCommand { get; private set; }
		public static string getUndoCommand { get; private set; }

		public static string getHotkeyAssignCommand { get; private set; }
		public static string getStartSessionCommand { get; private set; }

		public static string getDefineMobCommand { get; private set; }
		public static string getDefineItemCommand { get; private set; }

		public static string getCancelCommand { get; private set; }


		public enum Speech {
			NONE,
			START,
			STOP,
			PAUSE,
			SWITCH_GRAMMAR,
			CONFIRM,
			REFUSE,
			NEW_TARGET,
			UNDO,
			TARGET_KILLED,
			ASSIGN_HOTKEY_TO_ITEM,
			DEFINE_ITEM,
			DEFINE_MOB,
			CANCEL
		};

		public static Speech GetEnum(string text) {
			if (text == getStartCommand) {
				return Speech.START;
			}
			else if (text == getPauseCommand) {
				return Speech.PAUSE;
			}
			else if (text == getStopCommand) {
				return Speech.STOP;
			}
			else if (text == getConfirmationCommand) {
				return Speech.CONFIRM;
			}
			else if (text == getRefusalCommand) {
				return Speech.REFUSE;
			}
			else if (text == getNewTargetCommand) {
				return Speech.NEW_TARGET;
			}
			else if (text == getTargetKilledCommand) {
				return Speech.TARGET_KILLED;
			}
			else if (text == getSwitchGrammarCommand) {
				return Speech.SWITCH_GRAMMAR;
			}
			else if (text == getUndoCommand) {
				return Speech.UNDO;
			}
			else if (text == getHotkeyAssignCommand) {
				return Speech.ASSIGN_HOTKEY_TO_ITEM;
			}
			else if (text == getDefineMobCommand) {
				return Speech.DEFINE_MOB;
			}
			else if (text == getDefineItemCommand) {
				return Speech.DEFINE_ITEM;
			}
			else if (text == getCancelCommand) {
				return Speech.CANCEL;
			}

			else {
				return Speech.NONE;
			}
		}

		public static string GetSpeechString(Speech speech) {
			switch (speech) {
				case Speech.NONE: {
					return "";
				}
				case Speech.START: {
					return getStartCommand;
				}
				case Speech.STOP: {
					return getStopCommand;
				}
				case Speech.PAUSE: {
					return getPauseCommand;
				}
				case Speech.SWITCH_GRAMMAR: {
					return getSwitchGrammarCommand;
				}
				case Speech.CONFIRM: {
					return getConfirmationCommand;
				}
				case Speech.REFUSE: {
					return getRefusalCommand;
				}
				case Speech.NEW_TARGET: {
					return getNewTargetCommand;
				}
				case Speech.UNDO: {
					return getUndoCommand;
				}
				case Speech.TARGET_KILLED: {
					return getTargetKilledCommand;
				}
				case Speech.ASSIGN_HOTKEY_TO_ITEM: {
					return getHotkeyAssignCommand;
				}
				case Speech.DEFINE_MOB: {
					return getDefineMobCommand;
				}
				case Speech.DEFINE_ITEM: {
					return getDefineItemCommand;
				}
				case Speech.CANCEL: {
					return getCancelCommand;
				}
			}
			throw new NotSupportedException();
		}

		static CCommands() {
			if (!File.Exists(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + CONTROL_DEFINITION_NAME)) {
				throw new CustomException("Could not locate 'Control.definition' file! You have to redownload this application");
			}
			using (StreamReader sr = File.OpenText(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + CONTROL_DEFINITION_NAME)) {
				while (!sr.EndOfStream) {
					string line = sr.ReadLine();
					if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) {
						continue;
					}
					string[] split = line.Split(':');
					string modified = split[1].Remove(0, 1);
					modified = modified.Trim();
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
						case "UNDO": {
							getUndoCommand = modified;
							break;
						}
						case "HOTKEY_ASSIGN": {
							getHotkeyAssignCommand = modified;
							break;
						}
						case "DEFINE_MOB": {
							getDefineMobCommand = modified;
							break;
						}
						case "DEFINE_ITEM": {
							getDefineItemCommand = modified;
							break;
						}
						case "CANCEL": {
							getCancelCommand = modified;
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
