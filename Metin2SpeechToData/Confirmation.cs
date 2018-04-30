using System;
using System.Threading;
using System.Speech.Recognition;

namespace Metin2SpeechToData {
	class Confirmation {
		private static SpeechRecognitionEngine _confimer = new SpeechRecognitionEngine();
		private static ManualResetEventSlim evnt = new ManualResetEventSlim(false);

		private static readonly string[] _boolConfirmation = new string[2] { Program.controlCommands.getConfirmationCommand, Program.controlCommands.getRefusalCommand };

		private static bool _booleanResult;

		public static void Initialize() {
			_confimer.SetInputToDefaultAudioDevice();
		}

		public static bool AskForBooleanConfirmation(string question) {
			evnt.Reset();
			_confimer.UnloadAllGrammars();
			_confimer.LoadGrammar(new Grammar(new Choices(_boolConfirmation)));
			_confimer.RecognizeAsync(RecognizeMode.Single);
			_confimer.SpeechRecognized += Confimer_SpeechRecognized;
			Console.WriteLine(question);
			evnt.Wait();
			return _booleanResult;
		}

		private static void Confimer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			_confimer.SpeechRecognized -= Confimer_SpeechRecognized;
			if (e.Result.Text == _boolConfirmation[0]) {
				_booleanResult = true;
			}
			else {
				_booleanResult = false;
			}
			_confimer.RecognizeAsyncStop();
			evnt.Set();
		}
	}
}
