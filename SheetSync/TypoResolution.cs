using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SheetSync {
	class TypoResolution {
		private static ManualResetEventSlim evnt = new ManualResetEventSlim();
		private static HotKeyMapper m = new HotKeyMapper();
		private static Typos[] typos;

		public TypoResolution() {

		}

		private void ResolveTypos() {
			m.hotkeyOverriding = true;
			for (int j = 0; j < typos.Length; j++) {
				evnt.Reset();
				Console.WriteLine("Typo: " + typos[j].originalTypo);
				Console.Write("Alternatives: ");
				for (int i = 0; i < typos[j].alternatives.Length - 1; i++) {
					Console.Write("(" + (i + 1) + ")-" + typos[j].alternatives[i] + ", ");
					m.AssignToHotkey(Keys.D1 + i, i + 1, Resolve);
				}
				Console.WriteLine("(" + (typos[j].alternatives.Length) + ")-" + typos[j].alternatives[typos[j].alternatives.Length - 1]);
				m.AssignToHotkey(Keys.D1 + typos[j].alternatives.Length - 1, typos[j].alternatives.Length - 1, Resolve);
				evnt.Wait();
			}
		}

		private int currentIndex = 0;
		private void Resolve(int selected) {
			Console.WriteLine("Replaced '" + typos[currentIndex].originalTypo + "' with '" + typos[currentIndex].alternatives[selected] + "'");
			typos[currentIndex].sheet.SetValue(typos[currentIndex].location.Address, typos[currentIndex].alternatives[selected]);
			evnt.Set();
			currentIndex++;
		}
	}
}
