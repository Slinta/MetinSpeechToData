using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Recognition;
using System.Text;
using System.Threading.Tasks;

namespace Metin2SpeechToData {
	public class Undo {
		private readonly SpeechRecognitionEngine undoRecogniser;
		public static Undo instance;
		private bool continueExecution = true;
		public bool ContinueExecution {
			get {
				bool returnVar = continueExecution;
				continueExecution = true;
				return returnVar;
			}
			set {
				continueExecution = value;
			}
		}
		public int operationProgressCounter = 0;

		public LinkedList<SessionSheet.ItemMeta> itemInsertionList { get; set; }


		enum CurrentOperation {
			ItemDropped,
			TargetKilled,
			Defining,
		}

		public Undo() {

			undoRecogniser = new SpeechRecognitionEngine();
			undoRecogniser.SetInputToDefaultAudioDevice();
			undoRecogniser.SpeechRecognized += UndoRecognised;
			undoRecogniser.LoadGrammar(new Grammar(new Choices(CCommands.getUndoCommand)));
			undoRecogniser.RecognizeAsync(RecognizeMode.Multiple);
			instance = this;
			itemInsertionList = new LinkedList<SessionSheet.ItemMeta>();

		}

		

		public void UndoRecognised(object sender, SpeechRecognizedEventArgs e) {
			Console.WriteLine("Undo Happened, Confidence: " + e.Result.Confidence);
			//if (itemInsertionList.Count == 0) {
			//	Console.WriteLine("Nothing else to undo!");
			//	return;
			//}
			//SessionSheet.ItemMeta action = itemInsertionList.First.Value;
			//Console.WriteLine("Would remove " + action.itemBase.mainPronounciation);

			//bool resultUndo = Confirmation.AskForBooleanConfirmation("'Confirm'/'Refuse'?");
			//if (resultUndo) {
			//	itemInsertionList.RemoveFirst();
			//	Console.WriteLine("Removed " + action.itemBase.mainPronounciation + " from the stack");

			//}
			//else {
			//	Console.WriteLine("Undo refused!");
			//}

			//Console.WriteLine();
			CancelDefining();
		}

		public void AddItem(DefinitionParserData.Item item, string enemy, DateTime dropTime, int amount) {
			itemInsertionList.AddFirst(new SessionSheet.ItemMeta(item, enemy, dropTime, amount));
		}

		private void CancelDefining () {
			continueExecution = false;
			HotKeyMapper.AbortReadLine();
		}
	}	
}
