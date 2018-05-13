using System;
using System.Speech.Recognition;
using Metin2SpeechToData.Structures;


namespace Metin2SpeechToData {
	public class GameRecognizer: RecognitionBase {

		public static event Modifier OnModifierRecognized;


		public SpeechRecognitionHelper helper { get; }
		public EnemyHandling enemyHandling { get; }

		public GameRecognizer(): base() {
			enemyHandling = new EnemyHandling();
			helper = new SpeechRecognitionHelper(this);
			currentState = RecognitionState.INACTIVE;

			mainRecognizer.LoadGrammar(new Grammar(new Choices(Program.controlCommands.getNewTargetCommand)));
			mainRecognizer.LoadGrammar(new Grammar(new Choices(Program.controlCommands.getTargetKilledCommand)));
			mainRecognizer.LoadGrammar(new Grammar(new Choices(Program.controlCommands.getUndoCommand)));
			mainRecognizer.LoadGrammar(new Grammar(new Choices(Program.controlCommands.getRemoveTargetCommand)));
			
			getCurrentGrammars.Add(Program.controlCommands.getNewTargetCommand, 0);
			getCurrentGrammars.Add(Program.controlCommands.getTargetKilledCommand, 1);
			getCurrentGrammars.Add(Program.controlCommands.getUndoCommand, 2);
			getCurrentGrammars.Add(Program.controlCommands.getRemoveTargetCommand, 3);
		}

		public override void SwitchGrammar(string grammarID) {
			Grammar selected = DefinitionParser.instance.GetGrammar(grammarID);
			if (DefinitionParser.instance.currentMobGrammarFile != null) {
				enemyHandling.SwitchGrammar(grammarID);
			}
			mainRecognizer.LoadGrammar(selected);
			base.SwitchGrammar(grammarID);
		}


		public override void OnRecognitionStateChanged(object sender, RecognitionState state) {
			base.OnRecognitionStateChanged(sender, state);
			switch (state) {
				case RecognitionState.PAUSED: {
					if (currentState == RecognitionState.ACTIVE) {
						DefinitionParser.instance.hotkeyParser.SetKeysActiveState(false);
						Console.WriteLine("Pausing Recognition");
					}
					else {
						Console.WriteLine("Recognition not running!");
						return;
					}
					break;
				}
				case RecognitionState.STOPPED: {
					if (currentState == RecognitionState.PAUSED) {
						Program.mapper.FreeCustomHotkeys();
					}
					break;
				}
			}
			currentState = state;
		}


		protected override void SpeechRecognized(object sender, SpeechRecognizedArgs args) {
			if (SpeechRecognitionHelper.reverseModifierDict.ContainsKey(args.text)) {
				ModifierRecognized(this, args);
				return;
			}
			Console.WriteLine(args.text + " -- " + args.confidence);
			DefinitionParserData.Item item = DefinitionParser.instance.currentGrammarFile.GetItemEntry(DefinitionParser.instance.currentGrammarFile.GetMainPronounciation(args.text));
			enemyHandling.ItemDropped(item);
		}


		protected override void ModifierRecognized(object sender, SpeechRecognizedArgs args) {
			SpeechRecognitionHelper.ModifierWords modifier = SpeechRecognitionHelper.reverseModifierDict[args.text];

			PreModiferEvaluation(modifier);

			OnModifierRecognized?.Invoke(this, new ModiferRecognizedEventArgs() {
				modifier = modifier,
				triggeringItem = args.text,
			});

			PostModiferEvaluation(modifier);
		}


		protected override void PreModiferEvaluation(SpeechRecognitionHelper.ModifierWords current) {
			base.PreModiferEvaluation(current);
			switch (current) {
				case SpeechRecognitionHelper.ModifierWords.NEW_TARGET: {
					mainRecognizer.Grammars[primaryGrammarIndex].Enabled = false;
					mainRecognizer.Grammars[getCurrentGrammars[Program.controlCommands.getNewTargetCommand]].Enabled = false;
					break;
				}
				case SpeechRecognitionHelper.ModifierWords.UNDO: {
					mainRecognizer.Grammars[primaryGrammarIndex].Enabled = false;
					break;
				}
			}
		}

		protected override void PostModiferEvaluation(SpeechRecognitionHelper.ModifierWords current) {
			base.PostModiferEvaluation(current);
			switch (current) {
				case SpeechRecognitionHelper.ModifierWords.NEW_TARGET: {
					mainRecognizer.Grammars[primaryGrammarIndex].Enabled = true;
					mainRecognizer.Grammars[getCurrentGrammars[Program.controlCommands.getNewTargetCommand]].Enabled = true;
					break;
				}
				case SpeechRecognitionHelper.ModifierWords.UNDO: {
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
	}
}
