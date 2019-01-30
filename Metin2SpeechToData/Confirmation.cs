using System;
using System.Threading;
using System.Speech.Recognition;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Metin2SpeechToData {
	static class Confirmation {
		private static SpeechRecognitionEngine _confimer = new SpeechRecognitionEngine();
		private static ManualResetEventSlim evnt = new ManualResetEventSlim(false);

		private static string[] _boolConfirmation;

		private static bool _booleanResult;

		private static List<int> grammarsThatWereEnabledBefore = new List<int>();
		public static void Initialize() {
			_boolConfirmation = new string[2] { Program.controlCommands.getConfirmationCommand, Program.controlCommands.getRefusalCommand };
			_confimer.SetInputToDefaultAudioDevice();
			_confimer.LoadGrammar(new Grammar(new Choices(_boolConfirmation)));
		}

		public static async Task<bool> AskForBooleanConfirmation(string question) {
			evnt.Reset();
			_confimer.RecognizeAsync(RecognizeMode.Multiple);
			_confimer.SpeechRecognized += Confimer_SpeechRecognized;
			Console.WriteLine(question);
			await Task.Run(() => evnt.Wait());
			return _booleanResult;
		}

		private static void Confimer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			bool speechValidated = false;
			if (e.Result.Text == _boolConfirmation[0]) {
				_booleanResult = true;
				speechValidated = true;
			}
			else if (e.Result.Text == _boolConfirmation[1]) {
				_booleanResult = false;
				speechValidated = true;
			}
			if (speechValidated) {
				_confimer.SpeechRecognized -= Confimer_SpeechRecognized;
				_confimer.RecognizeAsyncStop();
				evnt.Set();
			}
		}
	}
}
