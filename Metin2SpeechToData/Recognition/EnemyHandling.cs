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

		private EnemyState _state;
		public EnemyState state {
			get {
				return _state;
			}
			set {
				if(value == EnemyState.NO_ENEMY) {
					Console.ForegroundColor = ConsoleColor.Gray;
				}
				else {
					Console.ForegroundColor = ConsoleColor.Green;
				}
				_state = value;
			}
		}

		private readonly SpeechRecognitionEngine masterMobRecognizer;
		private readonly ManualResetEventSlim evnt;
		public string currentEnemy { get; set; }

		private readonly GameRecognizer asociated;

		public EnemyHandling(GameRecognizer recognizer) {
			asociated = recognizer;
			recognizer.OnModifierRecognized += EnemyTargetingModifierRecognized;
			evnt = new ManualResetEventSlim(false);
			masterMobRecognizer = new SpeechRecognitionEngine();
			masterMobRecognizer.SetInputToDefaultAudioDevice();
			masterMobRecognizer.SpeechRecognized += MasterMobRecognizer_SpeechRecognized;
			masterMobRecognizer.LoadGrammar(new Grammar(new Choices(CCommands.getCancelCommand)));
			Undo.instance.SubscribeEnemyHandler(this);
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

		#region New enemy target recognition
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
		#endregion

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
						if (enemy == CCommands.getCancelCommand) {
							Console.WriteLine("Targetting cancelled!");
							return;
						}
						enemy = DefinitionParser.instance.currentMobGrammarFile.GetMainPronounciation(enemy);
						state = EnemyState.FIGHTING;
						currentEnemy = enemy;
						Undo.instance.AddNewTarget(enemy);

						Console.WriteLine("Acquired target: " + currentEnemy);
						Console.WriteLine();
						return;
					}
					case EnemyState.FIGHTING: {

						EnemyKilled();

						EnemyTargetingModifierRecognized(this, args);
						return;
					}
				}
			}
			else if (args.modifier == CCommands.Speech.TARGET_KILLED) {
				EnemyKilled();
			}
		}

		/// <summary>
		/// Function to kill the last enemy if not killed explicitly by the user
		/// </summary>
		public void ForceKill() {
			EnemyTargetingModifierRecognized(this, new ModiferRecognizedEventArgs(CCommands.Speech.TARGET_KILLED, ""));
		}

		/// <summary>
		/// Kill prints and clenup for next enemy
		/// </summary>
		private void EnemyKilled() {
			Console.WriteLine();
			Console.WriteLine("Killed " + currentEnemy + ", the death count increased");
			_state = EnemyState.NO_ENEMY;

			Undo.instance.EnemyKilled(currentEnemy);
			currentEnemy = "";
		}

		/// <summary>
		/// Increases number count to 'item' in current speadsheet
		/// </summary>
		public void ItemDropped(DefinitionParserData.Item item, int amount = 1) {
			Undo.instance.AddItem(item, currentEnemy, DateTime.Now, amount);
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
