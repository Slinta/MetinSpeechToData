using System;
using System.Collections.Generic;
using System.Speech.Recognition;

namespace Metin2SpeechToData {
	public class Undo : IDisposable {

		public enum OperationTypes {
			None,
			TargetFound,
			ItemDropped,
			TargetKilled,
			Defining,
			UnUndoable,
		}

		private SpeechRecognitionEngine undoRecogniser;
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

		/// <summary>
		/// History of all item additions
		/// </summary>
		public LinkedList<SessionSheet.ItemMeta> itemInsertionList { get; set; }

		/// <summary>
		/// History of all enemy encounters
		/// </summary>
		public LinkedList<Target> enemyList { get; set; }

		/// <summary>
		/// History of new item definitions
		/// </summary>
		public LinkedList<string> definedItemList { get; set; }

		/// <summary>
		/// List that holds sequence of event in the lists above in the order they were said
		/// </summary>
		public LinkedList<OperationTypes> operations { get; set; }


		private OperationTypes currentActiveOperation;

		public Undo() {
			instance = this;
			itemInsertionList = new LinkedList<SessionSheet.ItemMeta>();
			operations = new LinkedList<OperationTypes>();
			definedItemList = new LinkedList<string>();
			enemyList = new LinkedList<Target>();
		}

		/// <summary>
		/// Sets up main Undo recognizer
		/// </summary>
		public void Initialize() {
			undoRecogniser = new SpeechRecognitionEngine();
			undoRecogniser.SetInputToDefaultAudioDevice();
			undoRecogniser.SpeechRecognized += UndoRecognized;
			undoRecogniser.LoadGrammar(new Grammar(new Choices(CCommands.getUndoCommand, CCommands.getCancelCommand)));
			undoRecogniser.RecognizeAsync(RecognizeMode.Multiple);
		}

		/// <summary>
		/// Call to attach EnemyHandlig to Undo
		/// </summary>
		public void SubscribeEnemyHandler(EnemyHandling reference) {
			enemyHandling = reference;
		}

		/// <summary>
		/// Sets current operation to 'type'
		/// </summary>
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
							UndoAddItem();
						}
						else if (op == OperationTypes.TargetKilled) {
							UndoTargetKilled();
						}
						else if (op == OperationTypes.TargetFound) {
							UndoAddNewTarget();
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
		/// <summary>
		/// Adds new Target from 'enemy' to internal lists and records it into Undo history
		/// </summary>
		public void AddNewTarget(string enemy) {
			Target t = new Target(enemy, DateTime.Now, true);
			enemyList.AddLast(t);
			operations.AddLast(OperationTypes.TargetFound);
		}

		//TODO: add comment to this function, I changed the else branfo to not throw, is the if/else needed ?
		public void EnemyKilled(string enemy) {
			if (enemy == enemyList.Last.Value.name) {
				//Everything is fine
				enemyList.AddLast(new Target(enemy, DateTime.Now, false));
				operations.AddLast(OperationTypes.TargetKilled);
			}
			else {
				Console.WriteLine("Attempted to kill an unexpected enemy '{0}'", enemy);
			}
		}

		//TODO: add comment to this function, why is it AddFirst here and nowhere else ?
		public void AddItem(DefinitionParserData.Item item, string enemy, DateTime dropTime, int amount) {
			itemInsertionList.AddFirst(new SessionSheet.ItemMeta(item, enemy, dropTime, amount));
			operations.AddLast(OperationTypes.ItemDropped);
		}
		#endregion


		#region Undo behaviour

		/// <summary>
		/// Undo Items
		/// </summary>
		private void UndoAddItem() {
			if (itemInsertionList.Count == 0) {
				Console.WriteLine("Nothing else to undo!");
				return;
			}
			SessionSheet.ItemMeta action = itemInsertionList.First.Value;

			bool resultUndo = Confirmation.AskForBooleanConfirmation("Would remove " + action.itemBase.mainPronounciation);

			if (resultUndo) {
				itemInsertionList.RemoveFirst();
				Console.WriteLine("Removed " + action.itemBase.mainPronounciation + " from the stack");

			}
			else {
				Console.WriteLine("Undo refused!");
			}
			Console.WriteLine();
		}

		/// <summary>
		/// Undo killing enemy
		/// </summary>
		private void UndoTargetKilled() {
			enemyHandling.currentEnemy = enemyList.Last.Value.name;
			enemyHandling.state = EnemyHandling.EnemyState.FIGHTING;
			enemyList.RemoveLast();
			Console.WriteLine("Now again fighting " + enemyList.Last.Value.name);
		}

		/// <summary>
		/// Undo acquiring new target
		/// </summary>
		private void UndoAddNewTarget() {
			enemyHandling.currentEnemy = "";
			Console.WriteLine("Reset current target to 'None'");
			enemyHandling.state = EnemyHandling.EnemyState.NO_ENEMY;
			enemyList.RemoveLast();
		}
		#endregion

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
