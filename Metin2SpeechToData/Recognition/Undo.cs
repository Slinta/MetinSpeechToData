using System;
using System.Collections.Generic;
using System.Speech.Recognition;

namespace Metin2SpeechToData {
	public class Undo : IDisposable {
		private readonly SpeechRecognitionEngine undoRecogniser;
		public static Undo instance { get; private set; }

		private bool continueExecution = true;

		private EnemyHandling enemyHandling;

		public bool isInterrupted {
			get {
				if (!continueExecution) {
					continueExecution = true;
					return false;
				}
				else {
					return continueExecution;
				}
			}
			set {
				continueExecution = value;
			}
		}

		public LinkedList<SessionSheet.ItemMeta> itemInsertionList { get; set; }
		public LinkedList<Target> enemyList { get; set; }
		public LinkedList<string> definedItemList { get; set; }

		public LinkedList<OperationTypes> operations { get; set; }


		private OperationTypes currentActiveOperation;


		public enum OperationTypes {
			None,
			TargetFound,
			ItemDropped,
			TargetKilled,
			Defining,
			UnUndoable,
		}

		public Undo() {
			undoRecogniser = new SpeechRecognitionEngine();
			undoRecogniser.SetInputToDefaultAudioDevice();
			undoRecogniser.SpeechRecognized += UndoRecognized;
			undoRecogniser.LoadGrammar(new Grammar(new Choices(CCommands.getUndoCommand, CCommands.getCancelCommand)));
			undoRecogniser.RecognizeAsync(RecognizeMode.Multiple);
			instance = this;
			itemInsertionList = new LinkedList<SessionSheet.ItemMeta>();
			operations = new LinkedList<OperationTypes>();
			definedItemList = new LinkedList<string>();
			enemyList = new LinkedList<Target>();
		}

		public void SubscribeEnemyHandler(EnemyHandling reference) {
			enemyHandling = reference;
		}

		public void SetCurrentOperation(OperationTypes type) {
			currentActiveOperation = type;
		}

		public void UndoRecognized(object sender, SpeechRecognizedEventArgs e) {
			if (e.Result.Confidence > Configuration.acceptanceThreshold) {
				switch (CCommands.GetEnum(e.Result.Text)) {
					case CCommands.Speech.UNDO: {

						Console.WriteLine("Undo Happened, Confidence: " + e.Result.Confidence);

						if (operations.Count == 0) {
							Console.WriteLine("Nothing to undo!");
							break;
						}

						OperationTypes op = operations.Last.Value;
						operations.RemoveLast();
						if (op == OperationTypes.ItemDropped) {
							UndoOneItem();
						}
						else if (op == OperationTypes.TargetKilled) {
							UndoTargetKilled();
						}
						else if (op == OperationTypes.TargetFound) {
							UndoTargetFound();
						}
						break;
					}
					case CCommands.Speech.CANCEL: {
						if (currentActiveOperation == OperationTypes.Defining) {
							CancelDefining();
						}
						break;
					}
				}
			}
		}

		#region Canceling
		private void CancelDefining() {
			continueExecution = false;
			HotKeyMapper.AbortReadLine();
			currentActiveOperation = OperationTypes.None;
		}
		#endregion

		#region Entry reception
		public void EnemyFound(string enemy) {
			Target t = new Target(enemy, DateTime.Now, true);
			enemyList.AddLast(t);
			operations.AddLast(OperationTypes.TargetFound);
		}

		public void EnemyKilled(string enemy) {

			if (enemy == enemyList.Last.Value.name) {
				//Everything is fine
				enemyList.AddLast(new Target(enemy, DateTime.Now, false));
				operations.AddLast(OperationTypes.TargetKilled);
			}
			else {
				throw new CustomException("Enemy killed wasn't found earlier");
			}
		}

		public void AddItem(DefinitionParserData.Item item, string enemy, DateTime dropTime, int amount) {
			itemInsertionList.AddFirst(new SessionSheet.ItemMeta(item, enemy, dropTime, amount));
			operations.AddLast(OperationTypes.ItemDropped);
		}
		#endregion


		private void UndoOneItem() {
			if (itemInsertionList.Count == 0) {
				Console.WriteLine("Nothing else to undo!");
				return;
			}
			SessionSheet.ItemMeta action = itemInsertionList.First.Value;
			Console.WriteLine("Would remove " + action.itemBase.mainPronounciation);

			bool resultUndo = Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'?");
			if (resultUndo) {
				itemInsertionList.RemoveFirst();
				Console.WriteLine("Removed " + action.itemBase.mainPronounciation + " from the stack");

			}
			else {
				Console.WriteLine("Undo refused!");
			}
			Console.WriteLine();
		}

		private void UndoTargetKilled() {
			enemyHandling.currentEnemy = enemyList.Last.Value.name;
			enemyHandling.State = EnemyHandling.EnemyState.FIGHTING;
			enemyList.RemoveLast();
			Console.WriteLine("Now again fighting " + enemyList.Last.Value.name);
		}

		private void UndoTargetFound() {
			enemyHandling.currentEnemy = "";
			Console.WriteLine("Reset current target to 'None'");
			enemyHandling.State = EnemyHandling.EnemyState.NO_ENEMY;
			enemyList.RemoveLast();
		}

		public struct Target {
			public TargetStates state { get; }
			public List<DefinitionParserData.Item> droppedItems { get; }
			public string name { get; }
			public DateTime killTime { get; }

			public Target(string constructedName, DateTime KillTime, bool found) {
				name = constructedName;
				droppedItems = new List<DefinitionParserData.Item>();
				killTime = KillTime;
				state = (found) ? TargetStates.Found : TargetStates.Killed;
			}
		}

		public enum TargetStates {
			None,
			Killed,
			Found,
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				instance = null;
				undoRecogniser.Dispose();
				enemyHandling = null;
				disposedValue = true;
			}
		}

		~Undo() {
			// Do not change this code. Put cleanup code in Dispose(bool disposing) above.
			Dispose(false);
		}

		// This code added to correctly implement the disposable pattern.
		public void Dispose() {
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
