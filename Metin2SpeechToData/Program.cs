using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Recognition;

namespace Metin2SpeechToData {
	class Program {
		public enum RecognitionState {
			ERROR,
			CONTROL_RUNNING,
			DROP_LOGGER_RUNNING,
			SWITCHING,
		}

		public delegate void Recognition(RecognitionState state);

		private static SpeechRecognitionEngine game;
		private static DefinitionParser parser;
		private static SpeechRecognitionHelper helper;

		public static bool debug = false;

		static void Main(string[] args) {
			if(args.Length == 0) {
				debug = false;
			}
			else if (args[0] == "debug") {
				debug = true;
			}

			parser = new DefinitionParser();
			game = new SpeechRecognitionEngine();
			if (debug) {
				Console.WriteLine(game.Grammars.Count);
			}
			helper = new SpeechRecognitionHelper(ref game);
			helper.OnRecognitionChange += OnRecognitionChange;
			Console.ReadKey();
		}

		private static void OnRecognitionChange(RecognitionState state) {
			switch (state) {
				case RecognitionState.ERROR: {
					Console.WriteLine("Something went wrong");
					break;
				}
				case RecognitionState.CONTROL_RUNNING: {
					game.SpeechRecognized -= Game_SpeechRecognized;
					game.RecognizeAsyncStop();
					break;
				}
				case RecognitionState.DROP_LOGGER_RUNNING: {
					game.SetInputToDefaultAudioDevice();
					game.SpeechRecognized += Game_SpeechRecognized;
					game.RecognizeAsync(RecognizeMode.Multiple);
					break;
				}
				default: {
					break;
				}
			}
		}

		private static void Game_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			Console.WriteLine(e.Result.Text + " -- " + e.Result.Confidence);
		}
	}
}
