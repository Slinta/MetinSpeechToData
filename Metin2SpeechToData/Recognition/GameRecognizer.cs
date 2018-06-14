using System;
using System.Speech.Recognition;
using Metin2SpeechToData.Structures;


namespace Metin2SpeechToData {
	public class GameRecognizer : RecognitionBase {

		public SpeechRecognitionHelper helper { get; }
		public EnemyHandling enemyHandling { get; }

		public event Modifier OnModifierRecognized;

		public GameRecognizer() : base() {
			enemyHandling = new EnemyHandling(this);
			helper = new SpeechRecognitionHelper(this);
			currentState = RecognitionState.INACTIVE;

			mainRecognizer.LoadGrammar(new Grammar(new Choices(CCommands.getNewTargetCommand)));
			mainRecognizer.LoadGrammar(new Grammar(new Choices(CCommands.getTargetKilledCommand)));

			getCurrentGrammars.Add(CCommands.getNewTargetCommand, 0);
			getCurrentGrammars.Add(CCommands.getTargetKilledCommand, 1);
		}

		public override void SwitchGrammar(string grammarID) {
			Grammar selected = DefinitionParser.instance.GetGrammar(grammarID);
			if (DefinitionParser.instance.currentMobGrammarFile != null) {
				enemyHandling.SwitchGrammar(grammarID);
			}
			if (mainRecognizer.Grammars.Count != 0) {
				for (int i = mainRecognizer.Grammars.Count - 1; i >= 0; i--) {
					if (mainRecognizer.Grammars[i].Name == grammarID) {
						mainRecognizer.UnloadGrammar(mainRecognizer.Grammars[i]);
					}
				}
			}

			mainRecognizer.LoadGrammar(selected);
			base.SwitchGrammar(grammarID);
		}


		public override void OnRecognitionStateChanged(object sender, RecognitionState state) {
			base.OnRecognitionStateChanged(sender, state);
			switch (state) {
				case RecognitionState.PAUSED: {
					if (currentState == RecognitionState.ACTIVE) {
						Program.mapper.ToggleItemHotkeys(false);
						Console.WriteLine();
						Console.WriteLine("Pausing item and mob recognition");
						Console.WriteLine("Reenable with: '" + CCommands.getStartCommand + "'");
						Console.WriteLine("Other commands: '" + CCommands.getStopCommand + "', '" + CCommands.getSwitchGrammarCommand + "'.");

					}
					else {
						Console.WriteLine("Recognition not running!");
						return;
					}
					break;
				}
				case RecognitionState.STOPPED: {
					if (currentState == RecognitionState.PAUSED) {
						Program.mapper.FreeItemHotkeys();
					}
					break;
				}
			}
			currentState = state;
		}


		protected override void SpeechRecognized(object sender, SpeechRecognizedArgs args) {
			if (SpeechHelperBase.reverseModifierDict.ContainsKey(args.text)) {
				ModifierRecognized(this, args);
				return;
			}
			Console.WriteLine(args.text + " -- " + args.confidence);
			DefinitionParserData.Item item = DefinitionParser.instance.currentGrammarFile.GetItemEntry(DefinitionParser.instance.currentGrammarFile.GetMainPronounciation(args.text));
			enemyHandling.ItemDropped(item);
		}


		protected override void ModifierRecognized(object sender, SpeechRecognizedArgs args) {
			CCommands.Speech modifier = SpeechHelperBase.reverseModifierDict[args.text];

			PreModiferEvaluation(modifier);

			OnModifierRecognized?.Invoke(this, new ModiferRecognizedEventArgs(modifier, args.text));

			PostModiferEvaluation(modifier);
		}


		protected override void PreModiferEvaluation(CCommands.Speech current) {
			base.PreModiferEvaluation(current);
			switch (current) {
				case CCommands.Speech.NEW_TARGET: {
					mainRecognizer.Grammars[primaryGrammarIndex].Enabled = false;
					mainRecognizer.Grammars[getCurrentGrammars[CCommands.GetSpeechString(current)]].Enabled = false;
					break;
				}
				case CCommands.Speech.UNDO: {
					mainRecognizer.Grammars[primaryGrammarIndex].Enabled = false;
					break;
				}
			}
		}

		protected override void PostModiferEvaluation(CCommands.Speech current) {
			base.PostModiferEvaluation(current);
			switch (current) {
				case CCommands.Speech.NEW_TARGET: {
					mainRecognizer.Grammars[primaryGrammarIndex].Enabled = true;
					mainRecognizer.Grammars[getCurrentGrammars[CCommands.GetSpeechString(current)]].Enabled = true;
					break;
				}
				case CCommands.Speech.UNDO: {
					mainRecognizer.Grammars[primaryGrammarIndex].Enabled = true;
					break;
				}
			}
		}

		/// <summary>
		/// Get the index of main grammar in the recognizer
		/// </summary>
		public int primaryGrammarIndex {
			get {
				return getCurrentGrammars[DefinitionParser.instance.currentGrammarFile.ID];
			}
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2213:DisposableFieldsShouldBeDisposed")]
		protected override void Dispose(bool disposing) {
			enemyHandling.Dispose();
			helper.Dispose();
			base.Dispose(disposing);
		}
	}
}
