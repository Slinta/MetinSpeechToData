using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Recognition;


namespace Metin2SpeechToData {
	public class GameRecognizer: IDisposable {

		public delegate void Recognition(object sender, RecognitionState state);
		public delegate void Modifier(object sender, ModiferRecognizedEventArgs args);

		public static event Modifier OnModifierRecognized;

		public enum RecognitionState {
			INACTIVE,
			ACTIVE,
			SWITCHING,
			PAUSED,
			STOPPED,
		}

		public RecognitionState currentState { get; private set; }

		public SpeechRecognitionHelper helper;
		public EnemyHandling enemyHandling;

		/// <summary>
		/// Get currently loded grammars by name with indexes
		/// </summary>
		public Dictionary<string, int> getCurrentGrammars { get; } = new Dictionary<string, int>();

		private SpeechRecognitionEngine mainGameRecognizer;

		public GameRecognizer() {
			mainGameRecognizer = new SpeechRecognitionEngine();
			mainGameRecognizer.SetInputToDefaultAudioDevice();

			enemyHandling = new EnemyHandling();
			helper = new SpeechRecognitionHelper(this);
			currentState = RecognitionState.INACTIVE;
			helper.OnRecognitionChange += OnRecognitionChange;

			mainGameRecognizer.LoadGrammar(new Grammar(new Choices(Program.controlCommands.getNewTargetCommand)));
			mainGameRecognizer.LoadGrammar(new Grammar(new Choices(Program.controlCommands.getTargetKilledCommand)));
			mainGameRecognizer.LoadGrammar(new Grammar(new Choices(Program.controlCommands.getUndoCommand)));
			getCurrentGrammars.Add(Program.controlCommands.getNewTargetCommand, 0);
			getCurrentGrammars.Add(Program.controlCommands.getTargetKilledCommand, 1);
			getCurrentGrammars.Add(Program.controlCommands.getUndoCommand, 2);
		}

		public void SwitchGrammar(string grammarID) {
			Grammar selected = DefinitionParser.instance.GetGrammar(grammarID);
			if(DefinitionParser.instance.currentMobGrammarFile != null) {
				enemyHandling.SwitchGrammar(grammarID);
			}
			mainGameRecognizer.LoadGrammar(selected);

			for (int i = 0; i < mainGameRecognizer.Grammars.Count; i++) {
				if(mainGameRecognizer.Grammars[i].Name == grammarID) {
					getCurrentGrammars.Add(grammarID, i);
				}
			}
		}


		private void OnRecognitionChange(object sender, RecognitionState state) {
			switch (state) {
				case RecognitionState.INACTIVE: {
					if (Program.debug) {
						Console.WriteLine("Currently inactive");
					}
					mainGameRecognizer.SpeechRecognized -= Game_SpeechRecognized_Wrapper;
					mainGameRecognizer.RecognizeAsyncStop();
					break;
				}
				case RecognitionState.ACTIVE: {
					if (Program.debug) {
						Console.WriteLine("Currently active");
					}
					mainGameRecognizer.SetInputToDefaultAudioDevice();
					mainGameRecognizer.SpeechRecognized += Game_SpeechRecognized_Wrapper;
					mainGameRecognizer.RecognizeAsync(RecognizeMode.Multiple);
					break;
				}
				case RecognitionState.PAUSED: {
					if (Program.debug) {
						Console.WriteLine("Currently paused");
					}
					break;
				}
				case RecognitionState.STOPPED: {
					if (Program.debug) {
						Console.WriteLine("Currently stoped");
					}
					break;
				}
				case RecognitionState.SWITCHING: {
					if (Program.debug) {
						Console.WriteLine("Currenly switching");
					}
					break;
				}
			}
		}

		private void Game_ModifierRecognized(object sender, SpeechRecognizedArgs e) {
			SpeechRecognitionHelper.ModifierWords modifier = SpeechRecognitionHelper.reverseModifierDict[e.text];

			#region TODO Disable unwanted grammars
			switch (modifier) {
				case SpeechRecognitionHelper.ModifierWords.NEW_TARGET: {
					break;
				}
				case SpeechRecognitionHelper.ModifierWords.UNDO: {
					break;
				}
				case SpeechRecognitionHelper.ModifierWords.TARGET_KILLED: {
					break;
				}
			}
			#endregion

			OnModifierRecognized?.Invoke(this, new ModiferRecognizedEventArgs() {
				modifier = modifier,
				triggeringItem = e.text,
				triggeringEnemy = ""
			});
		}


		private void Game_SpeechRecognized_Wrapper(object sender, SpeechRecognizedEventArgs e) {
			Game_SpeechRecognized(sender, new SpeechRecognizedArgs(e.Result.Text, e.Result.Confidence));
		}
		private void Game_SpeechRecognized(object sender, SpeechRecognizedArgs e) {
			if (SpeechRecognitionHelper.reverseModifierDict.ContainsKey(e.text)) {
				Game_ModifierRecognized(this, e);
				return;
			}
			Console.WriteLine(e.text + " -- " + e.confidence);
			DefinitionParserData.Item item = DefinitionParser.instance.currentGrammarFile.GetItemEntry(DefinitionParser.instance.currentGrammarFile.GetMainPronounciation(e.text));
			enemyHandling.ItemDropped(item);
		}

		/// <summary>
		/// Get the index of main grammar in the recognizer
		/// </summary>
		public int primaryGrammarIndex {
			get {
				return getCurrentGrammars[DefinitionParser.instance.currentGrammarFile.ID];
			}
		}

		public void Dispose() {
			((IDisposable)mainGameRecognizer).Dispose();
		}
	}
}
