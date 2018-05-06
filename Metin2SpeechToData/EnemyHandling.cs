using System;
using System.Speech.Recognition;
using System.Threading;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public class EnemyHandling {
		public enum EnemyState {
			NO_ENEMY,
			FIGHTING
		}

		public EnemyState state { get; set; }
		private SpeechRecognitionEngine masterMobRecognizer;
		private ManualResetEventSlim evnt;
		public MobAsociatedDrops mobDrops;
		private DropOutStack<ItemInsertion> stack;
		private string currentEnemy = "";
		private string currentItem = "";


		public EnemyHandling() {
			GameRecognizer.OnModifierRecognized += EnemyTargetingModifierRecognized;
			mobDrops = new MobAsociatedDrops();
			stack = new DropOutStack<ItemInsertion>(Configuration.undoHistoryLength);
			evnt = new ManualResetEventSlim(false);
			masterMobRecognizer = new SpeechRecognitionEngine();
			masterMobRecognizer.SetInputToDefaultAudioDevice();
			masterMobRecognizer.SpeechRecognized += MasterMobRecognizer_SpeechRecognized;
		}

		/// <summary>
		/// Switch mob grammar for area
		/// </summary>
		/// <param name="grammarID"></param>
		public void SwitchGrammar(string grammarID) {
			Grammar selected = DefinitionParser.instance.GetMobGrammar(grammarID);
			masterMobRecognizer.LoadGrammar(selected);
		}

		private string GetEnemy() {
			Console.WriteLine("Listening for enemy...");
			masterMobRecognizer.RecognizeAsync(RecognizeMode.Multiple);
			evnt.Wait();
			return currentEnemy;
		}

		private void MasterMobRecognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			masterMobRecognizer.RecognizeAsyncStop();
			currentEnemy = e.Result.Text;
			evnt.Reset();
		}

		/// <summary>
		/// Event fired after a modifier word was said
		/// </summary>
		/// <param name="keyWord">"NEW_TARGET" // "UNDO" // "REMOVE_TARGET" // TARGET_KILLED</param>
		/// <param name="args">Always supply at least string.Empty as args!</param>
		public void EnemyTargetingModifierRecognized(object sender, ModiferRecognizedEventArgs args) {
			if (args.modifier == SpeechRecognitionHelper.ModifierWords.NEW_TARGET) {
				switch (state) {
					case EnemyState.NO_ENEMY: {
						string enemy = GetEnemy();
						string actualEnemyName = DefinitionParser.instance.currentMobGrammarFile.GetMainPronounciation(args.triggeringEnemy);
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
						if (args.triggeringEnemy != "") {
							EnemyTargetingModifierRecognized(SpeechRecognitionHelper.ModifierWords.NEW_TARGET, args);
						}
						break;
					}
				}
			}
			else if (args.modifier == SpeechRecognitionHelper.ModifierWords.TARGET_KILLED) {
				Program.interaction.OpenWorksheet(DefinitionParser.instance.currentGrammarFile.ID);
				EnemyTargetingModifierRecognized(this, args);
			}
			else if (args.modifier == SpeechRecognitionHelper.ModifierWords.UNDO) {
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
					if (Program.interaction.currentSheet.Cells[action.addr.Row, action.addr.Column].GetValue<int>() == 0) {
						string itemName = Program.interaction.currentSheet.Cells[action.addr.Row, action.addr.Column - 2].Value.ToString();
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
		public void ItemDropped(DefinitionParserData.Item item, int amount = 1) {
			if (!string.IsNullOrWhiteSpace(currentEnemy)){
				mobDrops.UpdateDrops(currentEnemy, DefinitionParser.instance.currentGrammarFile.GetItemEntry(item.mainPronounciation));
			}
			ExcelCellAddress address = Program.interaction.GetAddress(item.mainPronounciation);
			Program.interaction.AddNumberTo(address, amount);
			stack.Push(new ItemInsertion { addr = address, count = amount });
		}

		public void CleanUp() {
			state = EnemyState.NO_ENEMY;
			mobDrops = null;
			stack.Clear();
			GameRecognizer.OnModifierRecognized -= EnemyTargetingModifierRecognized;
		}

		private struct ItemInsertion {
			public ExcelCellAddress addr;
			public int count;
		}
	}
}
