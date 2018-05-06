using System;
using System.Threading;
using System.Speech.Recognition;
using System.Collections.Generic;

namespace Metin2SpeechToData {
	class Confirmation {
		private static SpeechRecognitionEngine _confimer = new SpeechRecognitionEngine();
		private static ManualResetEventSlim evnt = new ManualResetEventSlim(false);

		private static string[] _boolConfirmation;

		private static bool _booleanResult;

		private static List<int> grammarsThatWereEnabledBefore = new List<int>();
		public static void Initialize() {
			_boolConfirmation = new string[2] { Program.controlCommands.getConfirmationCommand, Program.controlCommands.getRefusalCommand };
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

		public static void SelectivelyDisableEnableGrammars(ref SpeechRecognitionEngine engine, bool disable) {
			if (disable) {
				for (int i = 0; i < engine.Grammars.Count; i++) {
					if (engine.Grammars[i].Enabled) {
						engine.Grammars[i].Enabled = false;
						grammarsThatWereEnabledBefore.Add(i);
					}
				}
			}
			else {
				for (int i = 0; i < grammarsThatWereEnabledBefore.Count; i++) {
					engine.Grammars[grammarsThatWereEnabledBefore[i]].Enabled = true;
				}
				grammarsThatWereEnabledBefore.Clear();
			}

		}
	}
}
