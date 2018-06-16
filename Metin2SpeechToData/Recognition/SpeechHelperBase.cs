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
			{ CCommands.Speech.TARGET_KILLED, CCommands.getTargetKilledCommand },
			{ CCommands.Speech.UNDO, CCommands.getUndoCommand },
			{ CCommands.Speech.ASSIGN_HOTKEY_TO_ITEM, CCommands.getHotkeyAssignCommand },
		};

		/// <summary>
		/// Reverse dictionary for converting spoken word to enum entries
		/// </summary>
		public static readonly IReadOnlyDictionary<string, CCommands.Speech> reverseModifierDict = new Dictionary<string, CCommands.Speech>() {
			{ CCommands.getNewTargetCommand , CCommands.Speech.NEW_TARGET },
			{ CCommands.getTargetKilledCommand, CCommands.Speech.TARGET_KILLED },
			{ CCommands.getUndoCommand, CCommands.Speech.UNDO },
			{ CCommands.getHotkeyAssignCommand, CCommands.Speech.ASSIGN_HOTKEY_TO_ITEM },
		};


		protected SpeechRecognitionEngine controlingRecognizer;
		protected Dictionary<string, (int index, bool isActive)> _currentGrammars;
		protected readonly ManualResetEventSlim evnt = new ManualResetEventSlim();
		protected readonly RecognitionBase baseRecognizer;

		protected SpeechHelperBase(RecognitionBase master) {
			baseRecognizer = master;
		}

		protected virtual void InitializeControl() {
			controlingRecognizer = new SpeechRecognitionEngine();
			_currentGrammars = new Dictionary<string, (int, bool)>();

			string startC = CCommands.getStartCommand;
			string pauseC = CCommands.getPauseCommand;
			string switchC = CCommands.getSwitchGrammarCommand;
			string quitC = CCommands.getStopCommand;
			string defineMob = CCommands.getDefineMobCommand;
			string defineItem = CCommands.getDefineItemCommand;

			controlingRecognizer.LoadGrammar(new Grammar(new Choices(startC)) { Name = startC, Enabled = false });
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(pauseC)) { Name = pauseC, Enabled = false });
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(switchC)) { Name = switchC, Enabled = true });
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(quitC)) { Name = quitC, Enabled = true });
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(defineMob)) { Name = defineMob, Enabled = false });
			controlingRecognizer.LoadGrammar(new Grammar(new Choices(defineItem)) { Name = defineItem, Enabled = false });

			_currentGrammars.Add(startC, (0, false));
			_currentGrammars.Add(pauseC, (1, false));
			_currentGrammars.Add(switchC, (2, true));
			_currentGrammars.Add(quitC, (3, true));
			_currentGrammars.Add(defineMob, (4, false));
			_currentGrammars.Add(defineItem, (5, false));

			controlingRecognizer.SetInputToDefaultAudioDevice();
			controlingRecognizer.SpeechRecognized += Control_SpeechRecognized_Wrapper;

			controlingRecognizer.RecognizeAsync(RecognizeMode.Multiple);
		}

		protected void Control_SpeechRecognized_Wrapper(object sender, SpeechRecognizedEventArgs args) {
			if (Configuration.acceptanceThreshold < args.Result.Confidence) {
				Control_SpeechRecognized(new SpeechRecognizedArgs(args.Result.Text, args.Result.Confidence));
			}
		}

		protected void Switch_WordRecognized_Wrapper(object sender, SpeechRecognizedEventArgs e) {
			if (Configuration.acceptanceThreshold < e.Result.Confidence) {
				Switch_WordRecognized(new SpeechRecognizedArgs(e.Result.Text, e.Result.Confidence));
			}
		}

		protected abstract void Control_SpeechRecognized(SpeechRecognizedArgs e);
		protected abstract void Switch_WordRecognized(SpeechRecognizedArgs e);

		/// <summary>
		/// Enables/Disables grammar 'name' in controling recognizer according to value in 'avtive'
		/// </summary>
		/// <param name="name"></param>
		/// <param name="active"></param>
		public void SetGrammarState(string name, bool active) {
			if (_currentGrammars.ContainsKey(name)) {
				controlingRecognizer.Grammars[_currentGrammars[name].index].Enabled = active;
				_currentGrammars[name] = (_currentGrammars[name].index, active);
			}
			else {
				throw new CustomException("Grammar name '" + name + "' not found!");
			}
		}

		/// <summary>
		/// Prevents Console.ReadLine() from Main from consuming lines meant for different prompt
		/// </summary>
		public void AcquireControl() {
			evnt.Wait();
			controlingRecognizer.SpeechRecognized -= Control_SpeechRecognized_Wrapper;
		}

		/// <summary>
		/// Unlocs the ManualResetEvent which causes the control to return to the main class (Program)
		/// </summary>
		protected void ReturnControl() {
			evnt.Set();
			evnt.Dispose();
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				if (disposing) {
					_currentGrammars.Clear();
				}
				controlingRecognizer.RecognizeAsyncStop();
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

	public sealed class ModiferRecognizedEventArgs : EventArgs {
		public ModiferRecognizedEventArgs(CCommands.Speech modifier, string itemTrigger) {
			this.modifier = modifier;
			this.itemTrigger = itemTrigger;
		}
		public CCommands.Speech modifier { get; }
		public string itemTrigger { get; }
	}
}
