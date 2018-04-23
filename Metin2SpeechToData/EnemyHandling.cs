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
		public string getCurrentEnemy { get; private set; } = "";


		public EnemyHandling() {
			if (Program.debug) {
				WrittenControl.OnModifierWordHear += EnemyTargetingModifierRecognized;
			}
			Program.OnModifierWordHear += EnemyTargetingModifierRecognized;
			mobDrops = new MobAsociatedDrops();
		}

		~EnemyHandling() {
			if (Program.debug) {
				WrittenControl.OnModifierWordHear -= EnemyTargetingModifierRecognized;
			}
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
						if (args[0] == "" && getCurrentEnemy == "") {
							return;
						}
						string actualEnemyName = DefinitionParser.instance.currentMobGrammarFile.GetMainPronounciation(args[0]);
						state = EnemyState.FIGHTING;
						Program.interaction.OpenWorksheet(actualEnemyName);
						getCurrentEnemy = actualEnemyName;
						Console.WriteLine("Acquired target: " + getCurrentEnemy);
						break;
					}
					case EnemyState.FIGHTING: {
						state = EnemyState.NO_ENEMY;
						Console.WriteLine("Killed " + getCurrentEnemy + ", the death count increased");
						Program.interaction.AddNumberTo(new ExcelCellAddress(1, 5), 1);
						getCurrentEnemy = "";
						if (args[0] != "") {
							EnemyTargetingModifierRecognized(SpeechRecognitionHelper.ModifierWords.NEW_TARGET, args[0]);
						}
						break;
					}
				}
			}
			else if (keyWord == SpeechRecognitionHelper.ModifierWords.REMOVE_TARGET) {
				Program.interaction.OpenWorksheet(DefinitionParser.instance.currentGrammarFile.ID);
				getCurrentEnemy = "";
				state = EnemyState.NO_ENEMY;
			}
			else if (keyWord == SpeechRecognitionHelper.ModifierWords.TARGET_KILLED) {
				Program.interaction.OpenWorksheet(DefinitionParser.instance.currentGrammarFile.ID);
				EnemyTargetingModifierRecognized(SpeechRecognitionHelper.ModifierWords.NEW_TARGET, "");
			}
			else if (keyWord == SpeechRecognitionHelper.ModifierWords.UNDO) {
				ItemInsertion action = stack.Pop();
				if(action.addr == null) {
					Console.WriteLine("Nothing else to undo!");
					return;
				}
				Program.interaction.AddNumberTo(action.addr, -action.count);
				string itemName = Program.interaction.currentSheet.Cells[action.addr.Row, action.addr.Column - 2].Value.ToString();
				Console.WriteLine("Undoing... removed " + action.count + " items from " + Program.interaction.currentSheet.Cells[action.addr.Row, action.addr.Column - 2].Value);
				mobDrops.RemoveItemEntry(getCurrentEnemy, itemName);
			}
		}

		/// <summary>
		/// Increases number count to 'item' in current speadsheet
		/// </summary>
		public void ItemDropped(string item, int amount = 1) {
			//TODO: for text controlled input the address from name is not working due to the grammar not being additive
			string mainPron = DefinitionParser.instance.currentGrammarFile.GetMainPronounciation(item);
			if (!string.IsNullOrWhiteSpace(getCurrentEnemy)){
				mobDrops.UpdateDrops(getCurrentEnemy, mainPron);
			}
			ExcelCellAddress address = Program.interaction.AddressFromName(mainPron);
			Program.interaction.AddNumberTo(address, amount);
			stack.Push(new ItemInsertion { addr = address, count = amount });
		}

		private struct ItemInsertion {
			public ExcelCellAddress addr;
			public int count;
		}
	}
}
