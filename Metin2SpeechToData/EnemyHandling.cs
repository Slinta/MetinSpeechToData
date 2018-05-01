using System;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public class EnemyHandling {
		public enum EnemyState {
			NO_ENEMY,
			FIGHTING
		}

		public EnemyState state;
		public MobAsociatedDrops mobDrops;
		private DropOutStack<ItemInsertion> stack = new DropOutStack<ItemInsertion>(5);
		private string currentEnemy = "";
		private string currentItem = "";

		public EnemyHandling() {
			Program.OnModifierWordHear += EnemyTargetingModifierRecognized;
			mobDrops = new MobAsociatedDrops();
		}

		~EnemyHandling() {
			Program.OnModifierWordHear -= EnemyTargetingModifierRecognized;
		}

		/// <summary>
		/// Event fired after a modifier word was said
		/// </summary>
		/// <param name="keyWord">"NEW_TARGET" // "UNDO" // "REMOVE_TARGET" // TARGET_KILLED</param>
		/// <param name="args">Always supply at least string.Empty as args!</param>
		public void EnemyTargetingModifierRecognized(SpeechRecognitionHelper.ModifierWords keyWord, params string[] args) {
			if (keyWord == SpeechRecognitionHelper.ModifierWords.NEW_TARGET) {
				switch (state) {
					case EnemyState.NO_ENEMY: {
						//Initial state
						if (args[0] == "" && currentEnemy == "") {
							return;
						}
						string actualEnemyName = DefinitionParser.instance.currentMobGrammarFile.GetMainPronounciation(args[0]);
						state = EnemyState.FIGHTING;
						Program.interaction.OpenWorksheet(actualEnemyName);
						currentEnemy = actualEnemyName;
						Console.WriteLine("Acquired target: " + currentEnemy);
						stack.Clear();
						break;
					}
					case EnemyState.FIGHTING: {
						state = EnemyState.NO_ENEMY;
						Console.WriteLine("Killed " + currentEnemy + ", the death count increased");
						Program.interaction.AddNumberTo(new ExcelCellAddress(1, 5), 1);
						currentEnemy = "";
						stack.Clear();
						if (args[0] != "") {
							EnemyTargetingModifierRecognized(SpeechRecognitionHelper.ModifierWords.NEW_TARGET, args[0]);
						}
						break;
					}
				}
			}
			else if (keyWord == SpeechRecognitionHelper.ModifierWords.TARGET_KILLED) {
				Program.interaction.OpenWorksheet(DefinitionParser.instance.currentGrammarFile.ID);
				EnemyTargetingModifierRecognized(SpeechRecognitionHelper.ModifierWords.NEW_TARGET, "");
			}
			else if (keyWord == SpeechRecognitionHelper.ModifierWords.UNDO) {
				ItemInsertion action = stack.Peek();
				if (action.addr == null) {
					Console.WriteLine("Nothing else to undo!");
					return;
				}
				Console.WriteLine("Would remove " + action.count + " items from " + Program.interaction.currentSheet.Cells[action.addr.Row, action.addr.Column - 2].Value);

				bool resultUndo = Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'?");
				if (resultUndo) {
					action = stack.Pop();
					Program.interaction.AddNumberTo(action.addr, -action.count);
					string itemName = Program.interaction.currentSheet.Cells[action.addr.Row, action.addr.Column - 2].Value.ToString();
					if (Program.interaction.currentSheet.Cells[action.addr.Row, action.addr.Column].GetValue<int>() == 0) {
						Console.WriteLine("Remove " + currentItem + " from current enemy's (" + currentEnemy + ") item list?");
						bool resultRemoveFromFile = Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'?");
						if (resultRemoveFromFile) {
							currentItem = itemName;
							mobDrops.RemoveItemEntry(currentEnemy, currentItem, true);
						}
					}
				}
			}
		}

		/// <summary>
		/// Increases number count to 'item' in current speadsheet
		/// </summary>
		public void ItemDropped(string item, int amount = 1) {
			string mainPronounciation = DefinitionParser.instance.currentGrammarFile.GetMainPronounciation(item);
			if (!string.IsNullOrWhiteSpace(currentEnemy)){
				mobDrops.UpdateDrops(currentEnemy, DefinitionParser.instance.currentGrammarFile.GetItemEntry(mainPronounciation));
			}
			ExcelCellAddress address = Program.interaction.GetAddress(mainPronounciation);
			Program.interaction.AddNumberTo(address, amount);
			stack.Push(new ItemInsertion { addr = address, count = amount });
		}

		private struct ItemInsertion {
			public ExcelCellAddress addr;
			public int count;
		}
	}
}
