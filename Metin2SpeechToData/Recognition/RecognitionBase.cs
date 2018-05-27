using System;
using System.Collections.Generic;
using System.Speech.Recognition;
using Metin2SpeechToData.Structures;

namespace Metin2SpeechToData {
	public abstract class RecognitionBase: IDisposable {
		/// <summary>
		/// Available recognizer states
		/// </summary>
		public enum RecognitionState {
			INACTIVE,
			STOPPED,
			SWITCHING,
			GRAMMAR_SELECTED,
			PAUSED,
			ACTIVE,
		}


		protected delegate void Recognition(object sender, RecognitionState state);
		public delegate void Modifier(object sender, ModiferRecognizedEventArgs e);

		protected Dictionary<string, int> _currentGrammars;


		public bool isPrimaryDefinitionLoaded { get; set; }
		public RecognitionState currentState { get; set; }

		/// <summary>
		/// Get currently loded grammars by name with indexes
		/// </summary>
		public Dictionary<string,int> getCurrentGrammars {
			get { return _currentGrammars; }
		}

		/// <summary>
		/// Main recognizer, hold all the grammars
		/// </summary>
		protected SpeechRecognitionEngine mainRecognizer;

		/// <summary>
		/// Create and configure the base recognition engine
		/// </summary>
		protected RecognitionBase() {
			mainRecognizer = new SpeechRecognitionEngine();
			mainRecognizer.SetInputToDefaultAudioDevice();
			mainRecognizer.InitialSilenceTimeout = new TimeSpan(500);
			_currentGrammars = new Dictionary<string, int>();
		}

		/// <summary>
		/// Updates 'getCurrentGrammars' with newly loaded grammar
		/// </summary>
		/// <param name="grammarID"></param>
		public virtual void SwitchGrammar(string grammarID) {
			for (int i = 0; i < mainRecognizer.Grammars.Count; i++) {
				if (mainRecognizer.Grammars[i].Name == grammarID) {
					getCurrentGrammars.Add(grammarID, i);
				}
			}
		}

		/// <summary>
		/// Base speech recognized event handler
		/// </summary>
		private void Main_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
			if (Configuration.acceptanceThreshold < e.Result.Confidence) {
				SpeechRecognized(sender, new SpeechRecognizedArgs(e.Result.Text, e.Result.Confidence));
			}
		}

		/// <summary>
		/// Define standard speech recognized behaviour, if the word is determined as a modifier, call 'ModifierRecognized' and return
		/// </summary>
		protected abstract void SpeechRecognized(object sender, SpeechRecognizedArgs args);

		/// <summary>
		/// Define modifier recognized behaviour
		/// </summary>
		protected abstract void ModifierRecognized(object sender, SpeechRecognizedArgs args);

		/// <summary>
		/// Handler for changing recognition state, 
		/// </summary>
		public virtual void OnRecognitionStateChanged(object sender, RecognitionState state) {
			switch (state) {
				case RecognitionState.INACTIVE: {
					if (Program.debug) {
						Console.WriteLine("Currently inactive");
					}
					StopRecognition(true);
					break;
				}
				case RecognitionState.ACTIVE: {
					if (Program.debug) {
						Console.WriteLine("Currently active");
					}
					if (currentState != RecognitionState.PAUSED) {
						BeginRecognition(true);
					}
					else {
						foreach (int index in getCurrentGrammars.Values) {
							mainRecognizer.Grammars[index].Enabled = true;
						}
					}
					break;
				}
				case RecognitionState.PAUSED: {
					if (Program.debug) {
						Console.WriteLine("Currently paused");
					}
					foreach (int index  in getCurrentGrammars.Values) {
						mainRecognizer.Grammars[index].Enabled = false;
					}
					break;
				}
				case RecognitionState.STOPPED: {
					if (Program.debug) {
						Console.WriteLine("Currently stoped");
					}
					mainRecognizer.UnloadAllGrammars();
					_currentGrammars.Clear();
					mainRecognizer.RecognizeAsyncStop();
					mainRecognizer.Dispose();
					break;
				}
				case RecognitionState.SWITCHING: {
					if (Program.debug) {
						Console.WriteLine("Currenly switching");
					}
					break;
				}
			}
		}

		/// <summary>
		/// If a modifier is recognized, you can modify current grammars and other things in this method implementation
		/// </summary>
		/// <param name="current">the modifier that will be switched to</param>
		protected virtual void PreModiferEvaluation(CCommands.Speech current) {
			if (Program.debug) {
				Console.WriteLine("Switching modifier to " + current);
			}
		}

		/// <summary>
		/// If a modifier is recognized, you can modify current grammars and other things in this method implementation
		/// </summary>
		/// <param name="current">the modifier that will be switched to</param>
		protected virtual void PostModiferEvaluation(CCommands.Speech current) {
			if (Program.debug) {
				Console.WriteLine("Modifier " + current + " handeled ");
			}
		}

		/// <summary>
		/// Sets up and start recognition
		/// </summary>
		public virtual void BeginRecognition(bool preformSetup = false) {
			if (preformSetup) {
				mainRecognizer.SetInputToDefaultAudioDevice();
				mainRecognizer.SpeechRecognized += Main_SpeechRecognized;
			}
			mainRecognizer.RecognizeAsync(RecognizeMode.Multiple);
		}

		/// <summary>
		/// Stops recognition
		/// </summary>
		public virtual void StopRecognition(bool preformCleanup = false) {
			if (preformCleanup) {
				mainRecognizer.SpeechRecognized -= Main_SpeechRecognized;
			}
			mainRecognizer.RecognizeAsyncStop();
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				if (disposing) {
					mainRecognizer.SpeechRecognized -= Main_SpeechRecognized;
				}
				mainRecognizer.Dispose();
				disposedValue = true;
			}
		}

		~RecognitionBase() {
			// Do not change this code. Put cleanup code in Dispose(bool disposing) above.
			Dispose(false);
		}

		// This code added to correctly implement the disposable pattern.
		public void Dispose() {
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
