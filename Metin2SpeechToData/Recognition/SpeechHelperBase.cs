using System;
using System.Collections.Generic;
using System.Speech.Recognition;
using System.Threading;
using Metin2SpeechToData.Structures;

namespace Metin2SpeechToData {
	public abstract class SpeechHelperBase : IDisposable {

		/// <summary>
		/// Modifiers dictionary, used to convert enum values to spoken word
		/// </summary>
		public static readonly IReadOnlyDictionary<CCommands.Speech, string> modifierDict = new Dictionary<CCommands.Speech, string>() {
			{ CCommands.Speech.NEW_TARGET , CCommands.getNewTargetCommand },
			{ CCommands.Speech.REMOVE_TARGET, CCommands.getRemoveTargetCommand },
			{ CCommands.Speech.TARGET_KILLED, CCommands.getTargetKilledCommand },
			{ CCommands.Speech.UNDO, CCommands.getUndoCommand },
			{ CCommands.Speech.ASSIGN_HOTKEY_TO_ITEM, CCommands.getHotkeyAssignCommand }
		};

		/// <summary>
		/// Reverse dictionary for converting spoken word to enum entries
		/// </summary>
		public static readonly IReadOnlyDictionary<string, CCommands.Speech> reverseModifierDict = new Dictionary<string, CCommands.Speech>() {
			{ CCommands.getNewTargetCommand , CCommands.Speech.NEW_TARGET },
			{ CCommands.getRemoveTargetCommand, CCommands.Speech.REMOVE_TARGET  },
			{ CCommands.getTargetKilledCommand, CCommands.Speech.TARGET_KILLED },
			{ CCommands.getUndoCommand, CCommands.Speech.UNDO },
			{ CCommands.getHotkeyAssignCommand, CCommands.Speech.ASSIGN_HOTKEY_TO_ITEM },
		};


		protected SpeechRecognitionEngine controlingRecognizer;
		protected Dictionary<string, (int index, bool isActive)> _currentGrammars;
		protected readonly ManualResetEventSlim evnt = new ManualResetEventSlim();

		protected virtual void InitializeControl() {
			controlingRecognizer = new SpeechRecognitionEngine();
			_currentGrammars = new Dictionary<string, (int, bool)>();

			string startC = CCommands.getStartCommand;
			string startSessionC = CCommands.getStartSessionCommand;
			string pauseC = CCommands.getPauseCommand;
			string switchC = CCommands.getSwitchGrammarCommand;
			string quitC = CCommands.getStopCommand;

			controlingRecognizer.LoadGrammar(new Grammar(new Choices(startC)) { Name = startC, Enabled = false });
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(startSessionC)) { Name = startSessionC, Enabled = false });
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(pauseC)) { Name = pauseC, Enabled = false });
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(switchC)) { Name = switchC, Enabled = true });
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(quitC)) { Name = quitC, Enabled = true });

			_currentGrammars.Add(startC, (0, false));
			_currentGrammars.Add(startSessionC, (1, false));
			_currentGrammars.Add(pauseC, (2, false));
			_currentGrammars.Add(switchC, (3, true));
			_currentGrammars.Add(quitC, (4, true));

			controlingRecognizer.SetInputToDefaultAudioDevice();
			controlingRecognizer.SpeechRecognized += Control_SpeechRecognized_Wrapper;

			controlingRecognizer.RecognizeAsync(RecognizeMode.Multiple);
		}

		protected void Control_SpeechRecognized_Wrapper(object sender, SpeechRecognizedEventArgs args) {
			if (Configuration.acceptanceThreshold < args.Result.Confidence) {
				Control_SpeechRecognized(new SpeechRecognizedArgs(args.Result.Text, args.Result.Confidence));
			}
		}

		protected void SwapReceiver(EventHandler<SpeechRecognizedEventArgs> unsub,
									EventHandler<SpeechRecognizedEventArgs> sub) {
			controlingRecognizer.SpeechRecognized -= unsub;
			controlingRecognizer.SpeechRecognized += sub;
		}

		protected abstract void Control_SpeechRecognized(SpeechRecognizedArgs e);


		/// <summary>
		/// Prevents Console.ReadLine() from Main from consuming lines meant for different prompt
		/// </summary>
		public void AcquireControl() {
			EventHandler<SpeechRecognizedEventArgs> waitCancellation = new EventHandler<SpeechRecognizedEventArgs>(
			(object o, SpeechRecognizedEventArgs e) => {
				if (e.Result.Text == CCommands.getStopCommand) {
					evnt.Set();
				}
			});

			controlingRecognizer.SpeechRecognized += waitCancellation;
			evnt.Wait();
			controlingRecognizer.RecognizeAsyncStop();
			controlingRecognizer.SpeechRecognized -= waitCancellation;
			controlingRecognizer.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
			controlingRecognizer.Dispose();
		}

		/// <summary>
		/// Unlocs the ManualResetEvent which causes the control to return to the main class (Program)
		/// </summary>
		protected void ReturnControl() {
			Dispose(true);
			evnt.Set();
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				if (disposing) {
					_currentGrammars.Clear();
				}
				evnt.Dispose();
				controlingRecognizer.Dispose();
				disposedValue = true;
			}
		}

		~SpeechHelperBase() {
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
