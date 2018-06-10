using System;
using System.Speech.Recognition;
using System.Threading;

namespace Metin2SpeechToData {
	public class EnemyHandling : IDisposable {
		public enum EnemyState {
			NO_ENEMY,
			FIGHTING
		}

		public EnemyState state { get; set; }
		private readonly SpeechRecognitionEngine masterMobRecognizer;
		private readonly ManualResetEventSlim evnt;
		private string currentEnemy = "";
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
						Console.WriteLine();
						Console.ForegroundColor = ConsoleColor.Green;
						return;
					}
					case EnemyState.FIGHTING: {
						state = EnemyState.NO_ENEMY;

						Console.WriteLine();
						Console.WriteLine("Killed " + currentEnemy + ", the death count increased");
						Console.ForegroundColor = ConsoleColor.White;
						Program.interaction.currentSession.EnemyKilled(currentEnemy, DateTime.Now);

						currentEnemy = "";
						EnemyTargetingModifierRecognized(this, args);
						return;
					}
				}
			}
			else if (args.modifier == CCommands.Speech.TARGET_KILLED) {
				Console.WriteLine();
				Console.WriteLine("Killed " + currentEnemy + ", the death count increased");
				Console.ForegroundColor = ConsoleColor.Gray;
				Program.interaction.currentSession.EnemyKilled(currentEnemy, DateTime.Now);
				EnemyTargetingModifierRecognized(this, new ModiferRecognizedEventArgs(CCommands.Speech.REMOVE_TARGET, ""));
			}
			else if (args.modifier == CCommands.Speech.REMOVE_TARGET) {
				currentEnemy = "";
				state = EnemyState.NO_ENEMY;

				Console.WriteLine("Reset current target to 'None'");
			}
			else if (args.modifier == CCommands.Speech.UNDO) {

				if (Program.interaction.currentSession.itemInsertionList.Count == 0) {
					Console.WriteLine("Nothing else to undo!");
					return;
				}
				Console.ForegroundColor = ConsoleColor.Red;
				SessionSheet.ItemMeta action = Program.interaction.currentSession.itemInsertionList.First.Value;
				Console.WriteLine("Would remove " + action.itemBase.mainPronounciation);

				bool resultUndo = Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'?");
				if (resultUndo) {
					Program.interaction.currentSession.itemInsertionList.RemoveFirst();
					Console.WriteLine("Removed " + action.itemBase.mainPronounciation + " from the stack");

				}
				else {
					Console.WriteLine("Undo refused!");
				}

				if (currentEnemy == "") {
					Console.ForegroundColor = ConsoleColor.Gray;
				}
				else {
					Console.ForegroundColor = ConsoleColor.Green;
				}
				Console.WriteLine();
			}
		}

		public void ForceKill() {
			EnemyTargetingModifierRecognized(this, new ModiferRecognizedEventArgs(CCommands.Speech.TARGET_KILLED, ""));
		}

		/// <summary>
		/// Increases number count to 'item' in current speadsheet
		/// </summary>
		public void ItemDropped(DefinitionParserData.Item item, int amount = 1) {
			Program.interaction.currentSession.Add(item, currentEnemy, DateTime.Now, amount);
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
