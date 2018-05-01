using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
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

		[DllImport("User32.Dll", EntryPoint = "PostMessageA")]
		private static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, int lParam);


		private Dictionary<Keys, ActionStashSpeechArgs> voiceHotkeys = new Dictionary<Keys, ActionStashSpeechArgs>();
		private Dictionary<Keys, ActionStashString> controlHotkeys = new Dictionary<Keys, ActionStashString>();

		public HotKeyMapper() {
			HotKeyManager.HotKeyPressed += new EventHandler<HotKeyEventArgs>(HotKeyManager_HotKeyPressed);
			Console.WriteLine("Initialized Key mapping!");
		}

		#region Add Hotkeys
		/// <summary>
		/// Assign hotkey 'selectedKey' to call function 'action' with 'arguments'
		/// </summary>
		public void AssignToHotkey(Keys selectedKey, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey)) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = KeyModifiers.None,
				_ungerID = HotKeyManager.RegisterHotKey(selectedKey, KeyModifiers.None)}
			);
		}
		/// <summary>
		/// Assign hotkey 'selectedKey' + a 'modifier' key to call function 'action' with 'arguments'
		/// </summary>
		public void AssignToHotkey(Keys selectedKey, KeyModifiers modifier, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey)) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = modifier,
				_ungerID = HotKeyManager.RegisterHotKey(selectedKey, modifier)}
			);
		}
		/// <summary>
		/// Assign hotkey 'selectedKey' + 'modifier' keys to call function 'action' with 'arguments'
		/// </summary>
		public void AssignToHotkey(Keys selectedKey, KeyModifiers modifier1, KeyModifiers modifier2, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey)) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = modifier1 | modifier2,
				_ungerID = HotKeyManager.RegisterHotKey(selectedKey, modifier1 | modifier2)}
			);
		}
		/// <summary>
		/// Assign hotkey 'selectedKey' to call function 'action' with 'arguments'
		/// </summary>
		public void AssignToHotkey(Keys hotkey, string command) {
			if (controlHotkeys.ContainsKey(hotkey)) {
				return;
				throw new CustomException(hotkey + " already mapped to " + controlHotkeys[hotkey] + "!");
			}
			controlHotkeys.Add(hotkey, new ActionStashString() {
				_action = AbortReadLine,
				_data = command,
				_keyModifier = KeyModifiers.None,
				_ungerID = HotKeyManager.RegisterHotKey(hotkey, KeyModifiers.None)
			}
			);
		}
		/// <summary>
		/// Assign hotkey 'selectedKey' + a 'modifier' key to call function 'action' with 'arguments'
		/// </summary>
		public void AssignToHotkey(Keys hotkey, KeyModifiers modifier, string command) {
			if (controlHotkeys.ContainsKey(hotkey)) {
				return;
				throw new CustomException(hotkey + " already mapped to " + controlHotkeys[hotkey] + "!");
			}
			controlHotkeys.Add(hotkey, new ActionStashString() {
				_action = AbortReadLine,
				_data = command,
				_keyModifier = modifier,
				_ungerID = HotKeyManager.RegisterHotKey(hotkey, modifier)
			}
			);
		}
		/// <summary>
		/// Assign hotkey 'selectedKey' + 'modifier' keys to call function 'action' with 'arguments'
		/// </summary>
		public void AssignToHotkey(Keys hotkey, KeyModifiers modifier1, KeyModifiers modifier2, string command) {
			if (controlHotkeys.ContainsKey(hotkey)) {
				return;
				throw new CustomException(hotkey + " already mapped to " + controlHotkeys[hotkey] + "!");
			}
			controlHotkeys.Add(hotkey, new ActionStashString() {
				_action = AbortReadLine,
				_data = command,
				_keyModifier = modifier1 | modifier2,
				_ungerID = HotKeyManager.RegisterHotKey(hotkey, modifier1 | modifier2)
			}
			);
		}

		#endregion


		#region Free dictionaries
		/// <summary>
		/// Removes all stored hotkeys
		/// </summary>
		public void FreeAll() {
			FreeControl();
			FreeGame();
		}

		/// <summary>
		/// Removes all control hotkeys
		/// </summary>
		public void FreeControl() {
			foreach (KeyValuePair<Keys,ActionStashString> item in controlHotkeys) {
				HotKeyManager.UnregisterHotKey(item.Value._ungerID);
			}
			controlHotkeys.Clear();
		}
		/// <summary>
		/// Removes all game hotkeys
		/// </summary>
		public void FreeGame() {
			foreach (KeyValuePair<Keys, ActionStashSpeechArgs> item in voiceHotkeys) {
				HotKeyManager.UnregisterHotKey(item.Value._ungerID);
			}
			voiceHotkeys.Clear();
		}

		/// <summary>
		/// Selectively remove a hotkey
		/// </summary>
		public void Free(Keys hotkey, bool debug = false) {
			if (voiceHotkeys.ContainsKey(hotkey)) {
				voiceHotkeys.Remove(hotkey);
			}
			else if (controlHotkeys.ContainsKey(hotkey)) {
				controlHotkeys.Remove(hotkey);
			}
			else {
				if (debug) { Console.WriteLine(hotkey + " not found in any list!"); }
			}
		}
		#endregion

		public void SetInactive(Keys key, bool state) {
			if (controlHotkeys.ContainsKey(key)) {
				ActionStashString stash = controlHotkeys[key];
				stash._isInactive = state;
				controlHotkeys[key] = stash;
			}
			if (voiceHotkeys.ContainsKey(key)) {
				ActionStashSpeechArgs stash = voiceHotkeys[key];
				stash._isInactive = state;
				voiceHotkeys[key] = stash;
			}
		}

		private void AbortReadLine(string command) {
			Program.currCommand = command;
			Thread.Sleep(250);
			PostMessage(Process.GetCurrentProcess().MainWindowHandle, 0x100, 0x0D, 0);
		}

		private void HotKeyManager_HotKeyPressed(object sender, HotKeyEventArgs e) {
			if (controlHotkeys.ContainsKey(e.Key)) {
				ActionStashString stash = controlHotkeys[e.Key];
				if (stash._keyModifier == e.Modifiers && !stash._isInactive) {
					stash._action.Invoke(stash._data);
				}
				else if (stash._isInactive) {
					Console.WriteLine("This command in currenly inaccessible");
				}
			}
			else if (voiceHotkeys.ContainsKey(e.Key)) {
				ActionStashSpeechArgs stash = voiceHotkeys[e.Key];
				if (stash._keyModifier == e.Modifiers && !stash._isInactive) {
					stash._action.Invoke(stash._data);
				}
				else if (stash._isInactive) {
					Console.WriteLine("This command in currenly inaccessible");
				}
			}
			else {
				Console.WriteLine("Hotkey for " + e.Key + " is not assigned");
			}
		}

		private struct ActionStashSpeechArgs {
			public KeyModifiers _keyModifier;
			public Action<SpeechRecognizedArgs> _action;
			public SpeechRecognizedArgs _data;
			public int _ungerID;
			public bool _isInactive;
		}

		private struct ActionStashString {
			public KeyModifiers _keyModifier;
			public Action<string> _action;
			public string _data;
			public int _ungerID;
			public bool _isInactive;
		}
	}
}
