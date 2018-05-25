using System;
using System.Speech.Recognition;
using System.Threading;
using OfficeOpenXml;
using Metin2SpeechToData.Structures;

namespace Metin2SpeechToData {
	public class EnemyHandling : IDisposable {
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
		private readonly GameRecognizer asociated;

		public EnemyHandling(GameRecognizer recognizer) {
			asociated = recognizer;
			recognizer.OnModifierRecognized += EnemyTargetingModifierRecognized;
			mobDrops = new MobAsociatedDrops();
			stack = new DropOutStack<ItemInsertion>(Configuration.undoHistoryLength);
			evnt = new ManualResetEventSlim(false);
			masterMobRecognizer = new SpeechRecognitionEngine();
			masterMobRecognizer.SetInputToDefaultAudioDevice();
			masterMobRecognizer.SpeechRecognized += MasterMobRecognizer_SpeechRecognized;
			masterMobRecognizer.LoadGrammar(new Grammar(new Choices(CCommands.getRemoveTargetCommand)));
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
			if (args.modifier == CCommands.Speech.NEW_TARGET) {
				switch (state) {
					case EnemyState.NO_ENEMY: {
						string enemy = GetEnemy();
						evnt.Reset();
						if (enemy == CCommands.getRemoveTargetCommand) {
							Console.WriteLine("Targetting cancelled!");
							return;
						}
						string actualEnemyName = DefinitionParser.instance.currentMobGrammarFile.GetMainPronounciation(enemy);
						state = EnemyState.FIGHTING;
						currentEnemy = actualEnemyName;
						Console.WriteLine("Acquired target: " + currentEnemy);
						stack.Clear();
						return;
					}
					case EnemyState.FIGHTING: {
						state = EnemyState.NO_ENEMY;
						Console.WriteLine("Killed " + currentEnemy + ", the death count increased");
						//Program.interaction.AddNumberTo(new ExcelCellAddress(1, 5), 1);
						//TODO: set in session file

						currentEnemy = "";
						stack.Clear();
						EnemyTargetingModifierRecognized(this, args);
						return;
					}
				}
			}
			else if (args.modifier == CCommands.Speech.TARGET_KILLED) {
				Console.WriteLine("Killed " + currentEnemy + ", the death count increased");
				//Program.interaction.AddNumberTo(new ExcelCellAddress(1, 5), 1);
				//TODO: set in session file
				EnemyTargetingModifierRecognized(this, new ModiferRecognizedEventArgs(CCommands.Speech.REMOVE_TARGET,""));
			}
			else if (args.modifier == CCommands.Speech.REMOVE_TARGET) {
				currentEnemy = "";
				currentItem = "";
				state = EnemyState.NO_ENEMY;
				stack.Clear();
				Console.WriteLine("Reset current target to 'None'");
			}
			else if (args.modifier == CCommands.Speech.UNDO) {
				ItemInsertion action = stack.Peek();
				if (action.address == null) {
					Console.WriteLine("Nothing else to undo!");
					return;
				}
				Console.WriteLine("Would remove " + action.count + " items from " + Program.interaction.currentSession.current.Cells[action.address.Row, action.address.Column - 2].Value);

				bool resultUndo = Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'?");
				if (resultUndo) {
					action = stack.Pop();
					Program.interaction.AddNumberTo(action.address, -action.count);
					if (Program.interaction.currentSession.current.Cells[action.address.Row, action.address.Column].GetValue<int>() == 0 && currentEnemy != "") {
						string itemName = Program.interaction.currentSession.current.Cells[action.address.Row, action.address.Column - 2].Value.ToString();
						Console.WriteLine("Remove " + currentItem + " from current  session?");
						bool resultRemoveFromFile = Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'?");
						if (resultRemoveFromFile) {
							currentItem = itemName;
							//TODO remove item from session
						}
						else {
							Console.WriteLine("Session data NOT modified!");
						}
					}
				}
				else {
					Console.WriteLine("Undo refused!");
				}
			}
		}

		/// <summary>
		/// Increases number count to 'item' in current speadsheet
		/// </summary>
		public void ItemDropped(DefinitionParserData.Item item, int amount = 1) {
			if (!string.IsNullOrWhiteSpace(currentEnemy)) {
				mobDrops.UpdateDrops(currentEnemy, item);
				//TODO redirect to SessionSheet
			}
			ExcelCellAddress address = Program.interaction.GetAddress(item.mainPronounciation);
			Program.interaction.AddNumberTo(address, amount);
			stack.Push(new ItemInsertion(address,amount));
		}
		public void ItemDropped(string item, int amount = 1) {
			ItemDropped(DefinitionParser.instance.currentGrammarFile.GetItemEntry(item), amount);
		}

		#region IDisposable Support
		private bool disposedValue = false;

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				if (disposing) {
					state = EnemyState.NO_ENEMY;
					mobDrops = null;
					stack.Clear();
					asociated.OnModifierRecognized -= EnemyTargetingModifierRecognized;
					masterMobRecognizer.SpeechRecognized -= MasterMobRecognizer_SpeechRecognized;
				}
				evnt.Dispose();
				masterMobRecognizer.Dispose();
				disposedValue = true;
			}
		}

		~EnemyHandling() {
			Dispose(false);
		}

		public void Dispose() {
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
