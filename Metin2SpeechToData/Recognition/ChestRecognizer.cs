using System;
using System.Speech.Recognition;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Threading;
using Metin2SpeechToData.Structures;


namespace Metin2SpeechToData {
	public class ChestRecognizer : RecognitionBase {
		private readonly SpeechRecognitionEngine numbers;

		private readonly ManualResetEventSlim evnt;

		public SpeechRecognitionHelper helper { get; }


		public ChestRecognizer() : base() {
			helper = new SpeechRecognitionHelper(this);
			evnt = new ManualResetEventSlim(false);
			numbers = new SpeechRecognitionEngine();
			numbers.SetInputToDefaultAudioDevice();
			numbers.SpeechRecognized += Numbers_SpeechRecognized;
			string[] strs = new string[201];
			for (int i = 0; i < strs.Length; i++) {
				strs[i] = i.ToString();
			}
			numbers.LoadGrammar(new Grammar(new Choices(strs)));
			mainRecognizer.LoadGrammar(new Grammar(new Choices(CCommands.getUndoCommand)));
			getCurrentGrammars.Add(CCommands.getUndoCommand, 2);
		}


		public override void SwitchGrammar(string grammarID) {
			Grammar selected = DefinitionParser.instance.GetGrammar(grammarID);
			mainRecognizer.LoadGrammar(selected);
			base.SwitchGrammar(grammarID);
		}

		protected override void SpeechRecognized(object sender, SpeechRecognizedArgs args) {
			if (SpeechHelperBase.reverseModifierDict.ContainsKey(args.text)) {
				ModifierRecognized(this, args);
				return;
			}
			Console.WriteLine(args.text + " -- " + args.confidence);
			StopRecognition();
			numbers.RecognizeAsync(RecognizeMode.Multiple);
			evnt.Wait();
			//Now we have an address and how many items they received
			Console.WriteLine("Parsed: " + _count);
			Undo.instance.AddItem(
				DefinitionParser.instance.currentGrammarFile.GetItemEntry(DefinitionParser.instance.currentGrammarFile.GetMainPronounciation(args.text)),
				"RNG",
				DateTime.Now,
				_count
			);
			evnt.Reset();
		}

		protected override void ModifierRecognized(object sender, SpeechRecognizedArgs args) {
			CCommands.Speech current = SpeechHelperBase.reverseModifierDict[args.text];
			switch (current) {
				case CCommands.Speech.UNDO: {
					SessionSheet.ItemMeta peeked = Undo.instance.itemInsertionList.Last.Value;
					if (peeked.dropTime == default(DateTime)) {
						Console.WriteLine("Nothing else to undo...");
						return;
					}
					Console.WriteLine("Undoing... " + peeked.itemBase.mainPronounciation + " with " + peeked.amount + " items");
					if (Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'")) {
						Console.Write("Confirmed");
						Undo.instance.itemInsertionList.RemoveLast();
					}
					else {
						Console.Write("Refused");
					}
					return;
				}
				default: {
					Console.WriteLine("Unsupported word " + args.text);
					return;
				}
			}
		}

		private int _count = 0;
		private void Numbers_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			if (int.TryParse(e.Result.Text, out _count)) {
				evnt.Set();
				numbers.RecognizeAsyncStop();
				BeginRecognition();
			}
		}

		protected override void Dispose(bool disposing) {
			numbers.SpeechRecognized -= Numbers_SpeechRecognized;
			numbers.RecognizeAsyncStop();
			numbers.Dispose();
			evnt.Dispose();
			helper.Dispose();
			base.Dispose(disposing);
		}
	}
}
