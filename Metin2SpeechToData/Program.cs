using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Recognition;

namespace Metin2SpeechToData {
	class Program {
		public enum RecognitionState {
			CONTROL_RUNNING,
			DROP_LOGGER_RUNNING,
		}
		public delegate void Recognition(RecognitionState state);
		private event Recognition OnRecognitionChange;


		static void Main(string[] args) {
			DefinitionParser df = new DefinitionParser();
			SpeechRecognitionEngine game = new SpeechRecognitionEngine();
			SpeechRecognitionHelper hepler = new SpeechRecognitionHelper(ref game);

			Console.ReadKey();


		}
	}
}
