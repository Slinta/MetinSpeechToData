using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;

namespace Metin2SpeechToData {

	[Flags]
	public enum KeyModifiers {
		None = 0,
		Alt = 1,
		Control = 2,
		Shift = 4,
		Windows = 8,
		NoRepeat = 0x4000
	}

	public static class HotKeyManager {
		private delegate void RegisterHotKeyDelegate(IntPtr hwnd, int id, uint modifiers, uint key);
		private delegate void UnRegisterHotKeyDelegate(IntPtr hwnd, int id);

		public static event EventHandler<HotKeyEventArgs> HotKeyPressed;

		private static volatile MessageWindow _wnd;
		private static volatile IntPtr _hwnd;
		private static ManualResetEvent _windowReadyEvent = new ManualResetEvent(false);

		private static int _id = 0;

		static HotKeyManager() {
			Thread messageLoop = new Thread(delegate () {
				Application.Run(new MessageWindow());
			});
			messageLoop.Name = "MessageLoopThread";
			messageLoop.IsBackground = true;
			messageLoop.Start();
		}


		public static int RegisterHotKey(Keys key, KeyModifiers modifiers) {
			_windowReadyEvent.WaitOne();
			int id = Interlocked.Increment(ref _id);
			_wnd.Invoke(new RegisterHotKeyDelegate(RegisterHotKeyInternal), _hwnd, id, (uint)modifiers, (uint)key);
			return id;
		}

		public static void UnregisterHotKey(int id) {
			_wnd.Invoke(new UnRegisterHotKeyDelegate(UnRegisterHotKeyInternal), _hwnd, id);
		}


		private static void RegisterHotKeyInternal(IntPtr hwnd, int id, uint modifiers, uint key) {
			RegisterHotKey(hwnd, id, modifiers, key);
		}

		private static void UnRegisterHotKeyInternal(IntPtr hwnd, int id) {
			UnregisterHotKey(_hwnd, id);
		}

		private static void OnHotKeyPressed(HotKeyEventArgs e) {
			HotKeyPressed?.Invoke(null, e);
		}


		private class MessageWindow : Form {
			private const int WM_HOTKEY = 0x312;

			public MessageWindow() {
				_wnd = this;
				_hwnd = this.Handle;
				_windowReadyEvent.Set();
			}

			protected override void WndProc(ref Message m) {
				if (m.Msg == WM_HOTKEY) {
					HotKeyEventArgs e = new HotKeyEventArgs(m.LParam);
					OnHotKeyPressed(e);
				}
				base.WndProc(ref m);
			}

			protected override void SetVisibleCore(bool value) {
				// Ensure the window never becomes visible
				base.SetVisibleCore(false);
			}
		}

		[DllImport("user32", SetLastError = true)]
		private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

		[DllImport("user32", SetLastError = true)]
		private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
	}

	public class HotKeyEventArgs : EventArgs {
		public readonly Keys Key;
		public readonly KeyModifiers Modifiers;

		public HotKeyEventArgs(Keys key, KeyModifiers modifiers) {
			Key = key;
			Modifiers = modifiers;
		}

		public HotKeyEventArgs(IntPtr hotKeyParam) {
			uint param = (uint)hotKeyParam.ToInt64();
			Key = (Keys)((param & 0xffff0000) >> 16);
			Modifiers = (KeyModifiers)(param & 0x0000ffff);
		}
	}


	public class HotKeyMapper {

		private Dictionary<Keys, ActionStash<>> voiceHotkeys = new Dictionary<Keys, ActionStash>();
		private Dictionary<Keys, FunctionStash> controlHotkeys = new Dictionary<Keys, FunctionStash>();

		private Action<SpeechRecognizedArgs> f1_action; private SpeechRecognizedArgs f1_args;

		public HotKeyMapper(Keys key, KeyModifiers modifier = KeyModifiers.None) {
			HotKeyManager.RegisterHotKey(key, modifier);
			HotKeyManager.HotKeyPressed += new EventHandler<HotKeyEventArgs>(HotKeyManager_HotKeyPressed);
			Console.WriteLine("Initialized Key mapping!");
		}

		#region Add Hotkeys
		public void AssignToHotkey(Keys selectedKey, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey)) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			voiceHotkeys.Add(selectedKey, new ActionStash() { _action = action, _data = arguments, _keyModifiers = new KeyModifiers[0] });
		}

		public void AssignToHotkey(Keys selectedKey, KeyModifiers modifier, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey)) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			voiceHotkeys.Add(selectedKey, new ActionStash() { _action = action, _data = arguments, _keyModifiers = new KeyModifiers[1] { modifier } });
		}

		public void AssignToHotkey(Keys selectedKey, KeyModifiers modifier1, KeyModifiers modifier2, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey)) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			voiceHotkeys.Add(selectedKey, new ActionStash() { _action = action, _data = arguments, _keyModifiers = new KeyModifiers[2] { modifier1, modifier2 } });
		}
		#endregion

		public void AssignToHotkey(Keys hotkey, string command) {
			controlHotkeys.Add(hotkey, new ActionStash() { _action = AbortReadLine, _data = command, _keyModifiers = new KeyModifiers[0] });
		}

		private void AbortReadLine(string command) {
			Program.currCommand = command;
			Thread.Sleep(250);
			PostMessage(Process.GetCurrentProcess().MainWindowHandle, WM_KEYDOWN, VK_RETURN, 0);
		}

		public void Free() {
			f1_action = null;
		}


		public void FreeAll() {
			voiceHotkeys.Clear();
			controlHotkeys.Clear();
		}

		public void Free(Keys hotkey) {
			if (voiceHotkeys.ContainsKey(hotkey)) {

			}
			if (controlHotkeys.ContainsKey(hotkey)) {

			}

			switch (hotkey) {
				case "F1": {
					f1_action = null;
					return;
				}
				default:
				throw new CustomException("Hotkey " + hotkey + " not defined");
			}
		}

		private void HotKeyManager_HotKeyPressed(object sender, HotKeyEventArgs e) {
			try {
				switch (e.Key) {
					case Keys.F1: {
						f1_action.Invoke(f1_args);
						return;
					}
					case Keys.F8: {
						if (e.Modifiers == (KeyModifiers.Alt | KeyModifiers.Control)) {
							Program.currCommand = keycodeCommands[wipe_args.arg1];
							Thread.Sleep(250);
							wipe_function.Invoke(wipe_args.ptr, wipe_args.message, 0x0D, wipe_args.arg2);
						}
						return;
					}
				}
			}
			catch {
				Console.WriteLine("Hotkey for " + e.Key + " is not assigned");
			}
		}

		private struct ActionStash<T> {
			public KeyModifiers[] _keyModifiers;
			public Action<T> _action;
			public SpeechRecognizedArgs _data;
		}

		[DllImport("User32.Dll", EntryPoint = "PostMessageA")]
		private static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, int lParam);
		const int VK_RETURN = 0x0D;
		const int WM_KEYDOWN = 0x100;
	}
}
