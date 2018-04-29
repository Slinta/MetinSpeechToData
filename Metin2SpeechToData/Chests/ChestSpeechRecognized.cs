using System;
using System.Threading.Tasks;
using System.Speech.Recognition;
using OfficeOpenXml;
using System.Threading;

namespace Metin2SpeechToData.Chests {
	class ChestSpeechRecognized {
		private SpeechRecognitionEngine game;
		private SpeechRecognitionEngine numbers;
		private ChestVoiceManager manager;
		private DropOutStack<ItemInsertion> stack;
		private ManualResetEventSlim evnt;

		private bool undoTriggered = false;

		public ChestSpeechRecognized(ChestVoiceManager manager) {
			this.manager = manager;
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
		}



		public void Subscribe(SpeechRecognitionEngine game) {
			this.game = game;
			game.SpeechRecognized += Game_SpeechRecognized;
		}

		public void Unsubscribe(SpeechRecognitionEngine game) {
			game.SpeechRecognized -= Game_SpeechRecognized;
		}


		private void Game_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			foreach (string s in SpeechRecognitionHelper.modifierDict.Values) {
				if (s == e.Result.Text) {
					switch (s) {
						case "Undo": {
							ItemInsertion peeked = stack.Peek();
							if(peeked.addr == null) {
								Console.WriteLine("Nothing else to undo...");
								return;
							}
							Console.WriteLine("Undoing... " + Program.interaction.currentSheet.Cells[peeked.addr.Row, peeked.addr.Column - 2].GetValue<string>() + " with " + peeked.count + " items");
							Console.WriteLine("'Confirm'/'Refuse'");
							undoTriggered = true;
							return;
						}
						case "Confirm": {
							if (undoTriggered) {
								Console.Write("Confirming");
								undoTriggered = false;
								ItemInsertion poped = stack.Pop();
								Program.interaction.AddNumberTo(poped.addr, -poped.count);
							}
							return;
						}
						case "Refuse": {
							if (undoTriggered) {
								Console.Write("Refusing");
								undoTriggered = false;
							}
							return;
						}
						default: {
							return;
						}
					}
				}
			}
			Console.WriteLine(e.Result.Text + " -- " + e.Result.Confidence);
			ExcelCellAddress address = Program.interaction.GetAddress(e.Result.Text);
			game.RecognizeAsyncStop();
			numbers.RecognizeAsync(RecognizeMode.Single);
			evnt.Wait();
			Console.WriteLine("Parsed: " + _count);
			evnt.Reset();
			//Now we have an address and how many items they received
			stack.Push(new ItemInsertion() { addr = address, count = _count });
			Program.interaction.AddNumberTo(address, _count);
		}

		private int _count = 0;
		private void Numbers_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			if (int.TryParse(e.Result.Text, out _count)) {
				evnt.Set();
				game.RecognizeAsync(RecognizeMode.Multiple);
				return;
			}
			throw new CustomException("This can never happen bacause the grammar is designed to only have numbers between 0-200 inclusive written as digits");
		}

		private struct ItemInsertion {
			public ExcelCellAddress addr;
			public int count;
		}
	}
}
