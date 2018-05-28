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


		public ChestRecognizer(): base() {
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
			//string mainPronounciation =  DefinitionParser.instance.currentGrammarFile.GetMainPronounciation(args.text);
			//ExcelCellAddress address = Program.interaction.GetAddress(mainPronounciation);
			StopRecognition();
			numbers.RecognizeAsync(RecognizeMode.Multiple);
			evnt.Wait();
			//Now we have an address and how many items they received
			Console.WriteLine("Parsed: " + _count);
			//stack.Push(new ItemInsertion(address,_count));
			//Program.interaction.AddNumberTo(address, _count);
			evnt.Reset();
		}

		protected override void ModifierRecognized(object sender, SpeechRecognizedArgs args) {
			CCommands.Speech current = SpeechHelperBase.reverseModifierDict[args.text];
			switch (current) {
				case CCommands.Speech.UNDO: {
					ItemInsertion peeked = stack.Peek();
					if (peeked.address == null) {
						Console.WriteLine("Nothing else to undo...");
						return;
					}
					Console.WriteLine("Undoing... " + Program.interaction.currentSession.current.Cells[peeked.address.Row, peeked.address.Column - 2].GetValue<string>() + " with " + peeked.count + " items");
					if (Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'")) {
						Console.Write("Confirming");
						ItemInsertion poped = stack.Pop();
						//Program.interaction.AddNumberTo(poped.address, -poped.count);
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
			if (int.TryParse(e.Result.Text, out _count)) {
				evnt.Set();
				numbers.RecognizeAsyncStop();
				BeginRecognition();
				return;
			}
			throw new CustomException("This can never happen bacause the grammar is designed to only have numbers between 0-200 inclusive written as digits");
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
