﻿using System;
using System.Speech.Recognition;
using OfficeOpenXml;
using System.Threading;
using Metin2SpeechToData.Structures;


namespace Metin2SpeechToData {
	public class ChestRecognizer : RecognitionBase {

		private readonly SpeechRecognitionEngine numbers;

		private readonly DropOutStack<ItemInsertion> stack;
		private readonly ManualResetEventSlim evnt;

		public SpeechRecognitionHelper helper { get; }


		public ChestRecognizer() : base() {
			helper = new SpeechRecognitionHelper(this);
			stack = new DropOutStack<ItemInsertion>(5);
			evnt = new ManualResetEventSlim(false);
			numbers = new SpeechRecognitionEngine();
			numbers.SetInputToDefaultAudioDevice();
			numbers.SpeechRecognized += Numbers_SpeechRecognized;
			string[] strs = new string[201];
			for (int i = 0; i < strs.Length; i++) {
				strs[i] = i.ToString();
			}
			numbers.LoadGrammar(new Grammar(new Choices(strs)));
			mainRecognizer.LoadGrammar(new Grammar(new Choices(Program.controlCommands.getUndoCommand)));
			getCurrentGrammars.Add(Program.controlCommands.getUndoCommand, 2);
		}


		public override void SwitchGrammar(string grammarID) {
			Grammar selected = DefinitionParser.instance.GetGrammar(grammarID);
			mainRecognizer.LoadGrammar(selected);
			base.SwitchGrammar(grammarID);
		}

		protected override void SpeechRecognized(object sender, SpeechRecognizedEventDetails args) {
			if (SpeechRecognitionHelper.reverseModifierDict.ContainsKey(args.text)) {
				ModifierRecognized(this, args);
				return;
			}
			Console.WriteLine(args.text + " -- " + args.confidence);
			string mainPronounciation = DefinitionParser.instance.currentGrammarFile.GetMainPronounciation(args.text);
			ExcelCellAddress address = Program.interaction.GetAddress(mainPronounciation);
			StopRecognition();
			numbers.RecognizeAsync(RecognizeMode.Multiple);
			evnt.Wait();
			//Now we have an address and how many items they received
			Console.WriteLine("Parsed: " + _count);
			stack.Push(new ItemInsertion(address, _count));
			Program.interaction.AddNumberTo(address, _count);
			evnt.Reset();
		}

		protected async override void ModifierRecognized(object sender, SpeechRecognizedEventDetails args) {
			SpeechRecognitionHelper.ModifierWords current = SpeechRecognitionHelper.reverseModifierDict[args.text];
			switch (current) {
				case SpeechRecognitionHelper.ModifierWords.UNDO: {
					ItemInsertion peeked = stack.Peek();
					if (peeked.address == null) {
						Console.WriteLine("Nothing else to undo...");
						return;
					}
					Console.WriteLine("Undoing... " + Program.interaction.currentSheet.Cells[peeked.address.Row, peeked.address.Column - 2].GetValue<string>() + " with " + peeked.count + " items");
					if (await Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'")) {
						Console.Write("Confirming");
						ItemInsertion poped = stack.Pop();
						Program.interaction.AddNumberTo(poped.address, -poped.count);
					}
					else {
						Console.Write("Refusing");
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
			_count = int.Parse(e.Result.Text);
			evnt.Set();
			numbers.RecognizeAsyncStop();
			BeginRecognition();
		}
	}
}
