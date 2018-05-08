using System;
using System.Speech.Recognition;
using System.Threading;
using static Metin2SpeechToData.SpeechRecognitionHelper;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public class EnemyHandling {
		public enum EnemyState {
			NO_ENEMY,
			FIGHTING
		}

		public EnemyState state { get; set; }
		private readonly SpeechRecognitionEngine masterMobRecognizer;
		private readonly ManualResetEventSlim evnt;
		public MobAsociatedDrops mobDrops { get; private set; }
		private readonly DropOutStack<ItemInsertion> stack;
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
			masterMobRecognizer.LoadGrammar(new Grammar(new Choices(Program.controlCommands.getRemoveTargetCommand)));
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
			evnt.Set();
		}

		/// <summary>
		/// Event fired after a modifier word was said
		/// </summary>
		/// <param name="keyWord">"NEW_TARGET" // "UNDO" // "REMOVE_TARGET" // TARGET_KILLED</param>
		/// <param name="args">Always supply at least string.Empty as args!</param>
		private void EnemyTargetingModifierRecognized(object sender, ModiferRecognizedEventArgs args) {
			if (args.modifier == ModifierWords.NEW_TARGET) {

				switch (state) {
					case EnemyState.NO_ENEMY: {
						string enemy = GetEnemy();
						evnt.Reset();
						if (enemy == Program.controlCommands.getRemoveTargetCommand) {
							Console.WriteLine("Targetting calcelled!");
							return;
						}
						string actualEnemyName = DefinitionParser.instance.currentMobGrammarFile.GetMainPronounciation(enemy);
						state = EnemyState.FIGHTING;
						Program.interaction.OpenWorksheet(actualEnemyName);
						currentEnemy = actualEnemyName;
						Console.WriteLine("Acquired target: " + currentEnemy);
						stack.Clear();
						return;
					}
					case EnemyState.FIGHTING: {
						state = EnemyState.NO_ENEMY;
						Console.WriteLine("Killed " + currentEnemy + ", the death count increased");
						Program.interaction.AddNumberTo(new ExcelCellAddress(1, 5), 1);
						currentEnemy = "";
						stack.Clear();
						EnemyTargetingModifierRecognized(this, args);
						return;
					}
				}
			}
			else if (args.modifier == ModifierWords.TARGET_KILLED) {
				Console.WriteLine("Killed " + currentEnemy + ", the death count increased");
				Program.interaction.AddNumberTo(new ExcelCellAddress(1, 5), 1);
				EnemyTargetingModifierRecognized(this, new ModiferRecognizedEventArgs() { modifier = ModifierWords.REMOVE_TARGET });
			}
			else if (args.modifier == ModifierWords.REMOVE_TARGET) {
				Program.interaction.OpenWorksheet(DefinitionParser.instance.currentGrammarFile.ID);
				currentEnemy = "";
				currentItem = "";
				state = EnemyState.NO_ENEMY;
				stack.Clear();
				Console.WriteLine("Reset current target to 'None', switching to " + DefinitionParser.instance.currentGrammarFile.ID + " sheet.");
			}
			else if (args.modifier == ModifierWords.UNDO) {
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
			if (!string.IsNullOrWhiteSpace(currentEnemy)) {
				mobDrops.UpdateDrops(currentEnemy, item);
			}
			ExcelCellAddress address = Program.interaction.GetAddress(item.mainPronounciation);
			Program.interaction.AddNumberTo(address, amount);
			stack.Push(new ItemInsertion { addr = address, count = amount });
		}
		public void ItemDropped(string item, int amount = 1) {
			ItemDropped(DefinitionParser.instance.currentGrammarFile.GetItemEntry(item), amount);
		}


		public void CleanUp() {
			state = EnemyState.NO_ENEMY;
			mobDrops = null;
			stack.Clear();
			GameRecognizer.OnModifierRecognized -= EnemyTargetingModifierRecognized;
			evnt.Dispose();
			masterMobRecognizer.SpeechRecognized -= MasterMobRecognizer_SpeechRecognized;
			masterMobRecognizer.Dispose();
		}

		private struct ItemInsertion {
			public ExcelCellAddress addr;
			public int count;
		}
	}
}
