using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;

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

		private static Action f1_action;
		private static Action f2_action;
		private static Action f3_action;
		private static Action f4_action;
		private static Action f5_action;
		private static Action f6_action;
		private static Action f7_action;
		private static Action f8_action;

		public HotKeyMapper() {
			HotKeyManager.RegisterHotKey(Keys.F1, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F2, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F3, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F4, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F5, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F6, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F7, KeyModifiers.None);
			HotKeyManager.RegisterHotKey(Keys.F8, KeyModifiers.None);
			HotKeyManager.HotKeyPressed += new EventHandler<HotKeyEventArgs>(HotKeyManager_HotKeyPressed);

			Console.WriteLine("Initialized F1-F8 Key mapping!");
		}

		public void AssignToHotkey(string hotkey, Action action) {
			switch (hotkey) {
				case "F1": {
					f1_action = action;
					return;
				}
				case "F2": {
					f2_action = action;
					return;
				}
				case "F3": {
					f3_action = action;
					return;
				}
				case "F4": {
					f4_action = action;
					return;
				}
				case "F5": {
					f5_action = action;
					return;
				}
				case "F6": {
					f6_action = action;
					return;
				}
				case "F7": {
					f7_action = action;
					return;
				}
				case "F8": {
					f8_action = action;
					return;
				}
				default:
				throw new CustomException("This hotkey is invalid, use f1 to f8");
			}
		}

		private static void HotKeyManager_HotKeyPressed(object sender, HotKeyEventArgs e) {
			switch (e.Key) {
				case Keys.F1:
				f1_action.Invoke();
				return;
				case Keys.F2:
				f2_action.Invoke();
				return;
				case Keys.F3:
				f3_action.Invoke();
				return;
				case Keys.F4:
				f4_action.Invoke();
				return;
				case Keys.F5:
				f5_action.Invoke();
				return;
				case Keys.F6:
				f6_action.Invoke();
				return;
				case Keys.F7:
				f7_action.Invoke();
				return;
				case Keys.F8:
				f8_action.Invoke();
				return;
			}
		}
	}
}
