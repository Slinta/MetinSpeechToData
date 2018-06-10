using System;
using System.Speech.Recognition;
using System.Threading;
using System.Collections.Generic;

namespace Metin2SpeechToData {
	public class EnemyHandling : IDisposable {
		public enum EnemyState {
			NO_ENEMY,
			FIGHTING
		}
		private EnemyState state;
		public EnemyState State {
			get {
				return state;
			}
			set {
				if(value == EnemyState.NO_ENEMY) {
					Console.ForegroundColor = ConsoleColor.Gray;
				}
				else {
					Console.ForegroundColor = ConsoleColor.Green;
				}
				state = value;
			}
		}
		private readonly SpeechRecognitionEngine masterMobRecognizer;
		private readonly ManualResetEventSlim evnt;
		private Target currentEnemy;

		private readonly GameRecognizer asociated;

		public EnemyHandling(GameRecognizer recognizer) {
			asociated = recognizer;
			recognizer.OnModifierRecognized += EnemyTargetingModifierRecognized;
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
			//Console.WriteLine("Current number of grammars in masterMobRecognizer:" + masterMobRecognizer.Grammars.Count);
			if(masterMobRecognizer.Grammars.Count != 0) {
				for (int i = masterMobRecognizer.Grammars.Count - 1; i >= 0; i--) {
					if (masterMobRecognizer.Grammars[i].Name == "Mob_" + grammarID) {
						masterMobRecognizer.UnloadGrammar(masterMobRecognizer.Grammars[i]);
					}
				}
			}
			
			masterMobRecognizer.LoadGrammar(selected);
			
			
		}

		private string GetEnemy() {
			Console.WriteLine("Listening for enemy...");
			masterMobRecognizer.RecognizeAsync(RecognizeMode.Multiple);
			evnt.Wait();
			return currentEnemy.name;
		}

		private void MasterMobRecognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			masterMobRecognizer.RecognizeAsyncStop();
			currentEnemy = new Target((e.Result.Text));
			evnt.Set();
		}

		/// <summary>
		/// Event fired after a modifier word was said
		/// </summary>
		/// <param name="keyWord">"NEW_TARGET" // "UNDO" // "REMOVE_TARGET" // TARGET_KILLED</param>
		/// <param name="args">Always supply at least string.Empty as args!</param>
		private void EnemyTargetingModifierRecognized(object sender, ModiferRecognizedEventArgs args) {
			if (args.modifier == CCommands.Speech.NEW_TARGET) {
				switch (State) {
					case EnemyState.NO_ENEMY: {
						string enemy = GetEnemy();
						evnt.Reset();
						if (enemy == CCommands.getRemoveTargetCommand) {
							Console.WriteLine("Targetting cancelled!");
							return;
						}

						State = EnemyState.FIGHTING;
						currentEnemy = new Target(enemy);

						Console.WriteLine("Acquired target: " + currentEnemy.name);
						Console.WriteLine();
						return;
					}
					case EnemyState.FIGHTING: {
						State = EnemyState.NO_ENEMY;

						Console.WriteLine();
						Console.WriteLine("Killed " + currentEnemy + ", the death count increased");
						Program.interaction.currentSession.EnemyKilled(currentEnemy.name, DateTime.Now);

						currentEnemy = new Target();
						EnemyTargetingModifierRecognized(this, args);
						return;
					}
				}
			}
			else if (args.modifier == CCommands.Speech.TARGET_KILLED) {
				Console.WriteLine();
				Console.WriteLine("Killed " + currentEnemy + ", the death count increased");
				Program.interaction.currentSession.EnemyKilled(currentEnemy.name, DateTime.Now);
				EnemyTargetingModifierRecognized(this, new ModiferRecognizedEventArgs(CCommands.Speech.REMOVE_TARGET, ""));
			}
			else if (args.modifier == CCommands.Speech.REMOVE_TARGET) {
				currentEnemy = new Target();
				State = EnemyState.NO_ENEMY;

				Console.WriteLine("Reset current target to 'None'");
			}
			//else if (args.modifier == CCommands.Speech.UNDO) {

			//	if (Program.interaction.currentSession.itemInsertionList.Count == 0) {
			//		Console.WriteLine("Nothing else to undo!");
			//		return;
			//	}
			//	Console.ForegroundColor = ConsoleColor.Red;
			//	SessionSheet.ItemMeta action = Program.interaction.currentSession.itemInsertionList.First.Value;
			//	Console.WriteLine("Would remove " + action.itemBase.mainPronounciation);

			//	bool resultUndo = Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'?");
			//	if (resultUndo) {
			//		Program.interaction.currentSession.itemInsertionList.RemoveFirst();
			//		Console.WriteLine("Removed " + action.itemBase.mainPronounciation + " from the stack");

			//	}
			//	else {
			//		Console.WriteLine("Undo refused!");
			//	}

			//	if (currentEnemy.name == "") {
			//		Console.ForegroundColor = ConsoleColor.Gray;
			//	}
			//	else {
			//		Console.ForegroundColor = ConsoleColor.Green;
			//	}
			//	Console.WriteLine();
			//}
		}

		public void ForceKill() {
			EnemyTargetingModifierRecognized(this, new ModiferRecognizedEventArgs(CCommands.Speech.TARGET_KILLED, ""));
		}

		/// <summary>
		/// Increases number count to 'item' in current speadsheet
		/// </summary>
		public void ItemDropped(DefinitionParserData.Item item, int amount = 1) {
			Undo.instance.AddItem(item, currentEnemy.name, DateTime.Now, amount);
		}
		public void ItemDropped(string item, int amount = 1) {
			ItemDropped(DefinitionParser.instance.currentGrammarFile.GetItemEntry(item), amount);
		}

		private struct Target {
			public List<DefinitionParserData.Item> droppedItems;
			public string name;

			public Target(string constructedName) {
				name = constructedName;
				droppedItems = new List<DefinitionParserData.Item>();
			}
		}

		#region IDisposable Support
		private bool disposedValue = false;

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				if (disposing) {
					State = EnemyState.NO_ENEMY;
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
