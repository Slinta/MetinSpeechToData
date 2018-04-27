using System;
using System.Speech.Recognition;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	class WrittenControl {

		private ControlSpeechCommands controlCommands;
		private EnemyHandling enemyHandling;

		public static bool stayInDebug = true;
		private bool artificialStart = false;
		public static event Program.ModifierTrigger OnModifierWordHear;


		public WrittenControl(ControlSpeechCommands controlCommands, ref EnemyHandling handling) {
			this.controlCommands = controlCommands;
			enemyHandling = handling;
			Control();
		}

		private void Control() {
			Console.WriteLine("ret - Return to normal mode");
			Console.WriteLine("Entered debug mode");
			stayInDebug = true;
			//TODO fix this class to work with new changes
			while (stayInDebug) {
				bool newCycle = true;
				string command = Console.ReadLine();
				if (Debug_ControlSpeechRecongized(command)) {
					continue;
				}
				foreach (string val in SpeechRecognitionHelper.modifierDict.Values) {
					if (val.ToLower() == command.ToLower()) {
						Game_SpeechRecognized(command);
						newCycle = false;
						break;
					}
				}
				if (!newCycle) {
					continue;
				}
				if (artificialStart) {
					Game_SpeechRecognized(command);
				}
			}
		}

		private bool Debug_ControlSpeechRecongized(string res) {
			if (res == controlCommands.getStartCommand) {
				Console.WriteLine("Starting Recognition. Current grammar: You should Know!");
				artificialStart = true;
			}
			else if (res == controlCommands.getStopCommand) {
				Console.WriteLine("Stopping Recognition!");
				Environment.Exit(0);
			}
			else if (res == controlCommands.getPauseCommand) {
				Console.WriteLine("Stopping item name reciever");
				artificialStart = false;
			}
			else if (res == controlCommands.getSwitchGrammarCommand) {
				Choices definitions = new Choices();
				Console.Write("Switching Grammar, available: ");
				string[] available = DefinitionParser.instance.getDefinitionNames;
				for (int i = 0; i < available.Length; i++) {
					definitions.Add(available[i]);
					if (i == available.Length - 1) {
						Console.Write(available[i]);
					}
					else {
						Console.Write(available[i] + ", ");
					}
				}
				Console.WriteLine();
				string selected = Console.ReadLine();
				Switch_WordRecognized(selected);
			}
			else if (res == "ret") {
				stayInDebug = false;
				Console.WriteLine("Exitting debug mode!");
			}
			else if (res == "clear") {
				Console.Clear();
			}
			else {
				return false;
			}
			return true;
		}

		private void Switch_WordRecognized(string selected) {
			Console.WriteLine("\nSelected - " + selected);
			if (!new System.Text.RegularExpressions.Regex(@"\w+\s(C|c)hest[+-]").IsMatch(selected)) {
				DefinitionParser.instance.currentMobGrammarFile = DefinitionParser.instance.GetMobDefinitionByName("Mob_" + selected);
			}
			DefinitionParser.instance.currentGrammarFile = DefinitionParser.instance.GetDefinitionByName(selected);
			Program.interaction.OpenWorksheet(selected);
		}


		private void Game_SpeechRecognized(string speechKappa) {
			foreach (string s in SpeechRecognitionHelper.modifierDict.Values) {
				if (s == speechKappa) {
					switch (s) {
						case "New Target": {
							OnModifierWordHear?.Invoke(SpeechRecognitionHelper.currentModifier, "");
							Console.WriteLine("Listening for enemy type");
							SpeechRecognitionHelper.currentModifier = SpeechRecognitionHelper.ModifierWords.NEW_TARGET;
							Console.WriteLine("Currentl in New Target modifier block, enter enemy");
							string str = Console.ReadLine();
							Game_ModifierRecognized(str);
							return;
						}
						case "Undo": {
							Console.Write("Undoing...");
							OnModifierWordHear?.Invoke(SpeechRecognitionHelper.ModifierWords.UNDO, "");
							return;
						}
						case "Remove Target": {
							Console.Write("Switching back to default NONE target");
							OnModifierWordHear?.Invoke(SpeechRecognitionHelper.ModifierWords.REMOVE_TARGET, "");
							return;
						}
					}
				}
			}
			ExcelWorksheet old = Program.interaction.currentSheet;
			enemyHandling.ItemDropped(speechKappa);
			Console.WriteLine(speechKappa + " -- 100% confident.");
		}

		private void Game_ModifierRecognized(string modifier) {
			OnModifierWordHear?.Invoke(SpeechRecognitionHelper.currentModifier, modifier);
			SpeechRecognitionHelper.currentModifier = SpeechRecognitionHelper.ModifierWords.NONE;
		}
	}
}
