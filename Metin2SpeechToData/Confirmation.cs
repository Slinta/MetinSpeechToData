﻿using System;
using System.Threading;
using System.Speech.Recognition;

namespace Metin2SpeechToData {
	public static class Confirmation {
		private static SpeechRecognitionEngine _confimer = new SpeechRecognitionEngine();
		private static ManualResetEventSlim evnt = new ManualResetEventSlim(false);

		private static string[] _boolConfirmation;

		private static bool _booleanResult;

		public static void Initialize() {
			_boolConfirmation = new string[2] { CCommands.getConfirmationCommand, CCommands.getRefusalCommand };
			_confimer.SetInputToDefaultAudioDevice();
		}

		public static bool AskForBooleanConfirmation(string question) {
			evnt.Reset();
			_confimer.UnloadAllGrammars();
			_confimer.LoadGrammar(new Grammar(new Choices(_boolConfirmation)));
			_confimer.RecognizeAsync(RecognizeMode.Multiple);
			_confimer.SpeechRecognized += Confimer_SpeechRecognized;
			Console.WriteLine(question);
			Console.WriteLine("'{0}', {1}", _boolConfirmation[0], _boolConfirmation[1]);
			evnt.Wait();
			return _booleanResult;
		}

		public static bool WrittenConfirmation(string question) {
			Console.WriteLine(question + "\ny/n");
			string line = Console.ReadLine();
			if(line == "y" || line == "yes" || line == "Y") {
				return true;
			}
			return false;
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
