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

		private readonly Dictionary<int, string> keycodeCommands = new Dictionary<int, string> {
			{ 8, "wipe" },
		};

		private Action<object, SpeecRecognizedArgs> f1_action; private SpeecRecognizedArgs f1_args;
		private Action<object, SpeecRecognizedArgs> f2_action; private SpeecRecognizedArgs f2_args;
		private Action<object, SpeecRecognizedArgs> f3_action; private SpeecRecognizedArgs f3_args;
		private Action<object, SpeecRecognizedArgs> f4_action; private SpeecRecognizedArgs f4_args;
		private Action<object, SpeecRecognizedArgs> f5_action; private SpeecRecognizedArgs f5_args;
		private Action<object, SpeecRecognizedArgs> f6_action; private SpeecRecognizedArgs f6_args;
		private Action<object, SpeecRecognizedArgs> f7_action; private SpeecRecognizedArgs f7_args;
		private Action<object, SpeecRecognizedArgs> f8_action; private SpeecRecognizedArgs f8_args;
		private Func<IntPtr, UInt32, int, int, bool> wipe_function; private PostMessageArgs wipe_args;

		public HotKeyMapper() {
			HotKeyManager.RegisterHotKey(Keys.F1, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F2, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F3, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F4, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F5, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F6, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F7, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F8, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F8, KeyModifiers.Control | KeyModifiers.Alt);
			HotKeyManager.HotKeyPressed += new EventHandler<HotKeyEventArgs>(HotKeyManager_HotKeyPressed);

			Console.WriteLine("Initialized F1-F8 Key mapping!");
		}

		public void AssignToHotkey(string hotkey, Action<object, SpeecRecognizedArgs> action, SpeecRecognizedArgs arguments) {
			switch (hotkey) {
				case "F1": {
					f1_action = action;
					f1_args = arguments;
					return;
				}
				case "F2": {
					f2_action = action;
					f2_args = arguments;
					return;
				}
				case "F3": {
					f3_action = action;
					f3_args = arguments;
					return;
				}
				case "F4": {
					f4_action = action;
					f4_args = arguments;
					return;
				}
				case "F5": {
					f5_action = action;
					f5_args = arguments;
					return;
				}
				case "F6": {
					f6_action = action;
					f6_args = arguments;
					return;
				}
				case "F7": {
					f7_action = action;
					f7_args = arguments;
					return;
				}
				case "F8": {
					f8_action = action;
					f8_args = arguments;
					return;
				}
				default:
				throw new CustomException("This hotkey is invalid, use f1 to f8");
			}
		}

		public void AssignToHotkey(string hotkey, Func<IntPtr, UInt32, int, int, bool> function, PostMessageArgs arguments) {
			switch (hotkey) {
				case "F8_Alt_Ctrl": {
					wipe_function = function;
					wipe_args = arguments;
					return;
				}
			}
		}

		public void Free() {
			f1_action = null;
			f2_action = null;
			f3_action = null;
			f4_action = null;
			f5_action = null;
			f6_action = null;
			f7_action = null;
			f8_action = null;
		}

		public void Free(string hotkey) {
			switch (hotkey) {
				case "F1": {
					f1_action = null;
					return;
				}
				case "F2": {
					f2_action = null;
					return;
				}
				case "F3": {
					f3_action = null;
					return;
				}
				case "F4": {
					f4_action = null;
					return;
				}
				case "F5": {
					f5_action = null;
					return;
				}
				case "F6": {
					f6_action = null;
					return;
				}
				case "F7": {
					f7_action = null;
					return;
				}
				case "F8": {
					f8_action = null;
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
						f1_action.Invoke(this, f1_args);
						return;
					}
					case Keys.F2: {
						f2_action.Invoke(this, f2_args);
						return;
					}
					case Keys.F3: {
						f3_action.Invoke(this, f3_args);
						return;
					}
					case Keys.F4: {
						f4_action.Invoke(this, f4_args);
						return;
					}
					case Keys.F5: {
						f5_action.Invoke(this, f5_args);
						return;
					}
					case Keys.F6: {
						f6_action.Invoke(this, f6_args);
						return;
					}
					case Keys.F7: {
						f7_action.Invoke(this, f7_args);
						return;
					}
					case Keys.F8: {
						if (e.Modifiers == KeyModifiers.None) {
							f8_action.Invoke(this, f8_args);
						}
						else if (e.Modifiers == (KeyModifiers.Alt | KeyModifiers.Control)) {
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
	}
}
