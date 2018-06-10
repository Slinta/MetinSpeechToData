using Metin2SpeechToData;
using System;
using System.Threading;
using System.Windows.Forms;
using static SheetSync.Structures;

namespace SheetSync {
	internal class TypoResolution : IDisposable {
		private readonly ManualResetEventSlim evnt = new ManualResetEventSlim();
		private readonly HotKeyMapper m = new HotKeyMapper();

		public TypoResolution(Typos[] typos) {
			if (typos.Length > 0 && Confirmation.WrittenConfirmation("Resovle name typos?")) {
				ResolveTypos();
			}
			getTypos = typos;
		}

		public Typos[] getTypos { get; }

		private void ResolveTypos() {
			m.hotkeyOverriding = true;
			for (int j = 0; j < getTypos.Length; j++) {
				evnt.Reset();
				Console.WriteLine("Typo: " + getTypos[j].originalTypo);
				Console.Write("Alternatives: ");
				for (int i = 0; i < getTypos[j].alternatives.Length - 1; i++) {
					Console.Write("(" + (i + 1) + ")-" + getTypos[j].alternatives[i] + ", ");
					m.AssignToHotkey(Keys.D1 + i, i + 1, Resolve);
				}
				Console.WriteLine("(" + (getTypos[j].alternatives.Length) + ")-" + getTypos[j].alternatives[getTypos[j].alternatives.Length - 1]);
				m.AssignToHotkey(Keys.D1 + getTypos[j].alternatives.Length - 1, getTypos[j].alternatives.Length - 1, Resolve);
				evnt.Wait();
			}
		}

		private int currentIndex = 0;
		private void Resolve(int selected) {
			Console.WriteLine("Replaced '" + getTypos[currentIndex].originalTypo + "' with '" + getTypos[currentIndex].alternatives[selected] + "'");
			getTypos[currentIndex].sheet.SetValue(getTypos[currentIndex].location.Address, getTypos[currentIndex].alternatives[selected]);
			evnt.Set();
			currentIndex++;
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				evnt.Dispose();
				disposedValue = true;
			}
		}

		~TypoResolution() {
			// Do not change this code. Put cleanup code in Dispose(bool disposing) above.
			Dispose(false);
		}

		// This code added to correctly implement the disposable pattern.
		public void Dispose() {
			Dispose(true);
		}
		#endregion
	}
}
