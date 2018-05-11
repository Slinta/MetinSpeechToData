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


		private static void RegisterHotKeyInternal(IntPtr windowHandle, int id, uint modifiers, uint key) {
			RegisterHotKey(windowHandle, id, modifiers, key);
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


		private readonly Dictionary<Keys, ActionStashSpeechArgs> voiceHotkeys = new Dictionary<Keys, ActionStashSpeechArgs>();

		private readonly Dictionary<Keys, ActionStashString> controlHotkeys = new Dictionary<Keys, ActionStashString>();

		public HotKeyMapper() {
			HotKeyManager.HotKeyPressed += new EventHandler<HotKeyEventArgs>(HotKeyManager_HotKeyPressed);
		}

		#region Add Hotkeys
		/// <summary>
		/// Assign hotkey 'selectedKey' to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey)) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = KeyModifiers.None,
				_unregID = HotKeyManager.RegisterHotKey(selectedKey, KeyModifiers.None)
			}
			);
			return voiceHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' + a 'modifier' key to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, KeyModifiers modifier, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey)) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = modifier,
				_unregID = HotKeyManager.RegisterHotKey(selectedKey, modifier)
			}
			);
			return voiceHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' + 'modifier' keys to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, KeyModifiers modifier1, KeyModifiers modifier2, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey)) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = modifier1 | modifier2,
				_unregID = HotKeyManager.RegisterHotKey(selectedKey, modifier1 | modifier2)
			}
			);
			return voiceHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' + 'modifier' keys to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, KeyModifiers modifier1, KeyModifiers modifier2, KeyModifiers modifier3, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey)) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = modifier1 | modifier2 | modifier3,
				_unregID = HotKeyManager.RegisterHotKey(selectedKey, modifier1 | modifier2 | modifier3)
			}
			);
			return voiceHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys hotkey, string command) {
			if (controlHotkeys.ContainsKey(hotkey)) {
				throw new CustomException(hotkey + " already mapped to " + controlHotkeys[hotkey] + "!");
			}
			controlHotkeys.Add(hotkey, new ActionStashString() {
				_action = AbortReadLine,
				_data = command,
				_keyModifier = KeyModifiers.None,
				_unregID = HotKeyManager.RegisterHotKey(hotkey, KeyModifiers.None)
			}
			);
			return controlHotkeys[hotkey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' + a 'modifier' key to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys hotkey, KeyModifiers modifier, string command) {
			if (controlHotkeys.ContainsKey(hotkey)) {
				throw new CustomException(hotkey + " already mapped to " + controlHotkeys[hotkey] + "!");
			}
			controlHotkeys.Add(hotkey, new ActionStashString() {
				_action = AbortReadLine,
				_data = command,
				_keyModifier = modifier,
				_unregID = HotKeyManager.RegisterHotKey(hotkey, modifier)
			}
			);
			return controlHotkeys[hotkey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' + 'modifier' keys to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys hotkey, KeyModifiers modifier1, KeyModifiers modifier2, string command) {
			if (controlHotkeys.ContainsKey(hotkey)) {
				throw new CustomException(hotkey + " already mapped to " + controlHotkeys[hotkey] + "!");
			}
			controlHotkeys.Add(hotkey, new ActionStashString() {
				_action = AbortReadLine,
				_data = command,
				_keyModifier = modifier1 | modifier2,
				_unregID = HotKeyManager.RegisterHotKey(hotkey, modifier1 | modifier2)
			}
			);
			return controlHotkeys[hotkey]._unregID;
		}

		#endregion


		#region Free dictionaries
		/// <summary>
		/// Removes all stored hotkeys
		/// </summary>
		public void FreeAllHotkeys() {
			FreeControlHotkeys();
			FreeGameHotkeys();
		}

		/// <summary>
		/// Removes all control hotkeys
		/// </summary>
		public void FreeControlHotkeys() {
			foreach (KeyValuePair<Keys, ActionStashString> item in controlHotkeys) {
				HotKeyManager.UnregisterHotKey(item.Value._unregID);
			}
			controlHotkeys.Clear();
		}

		/// <summary>
		/// Removes all game hotkeys
		/// </summary>
		public void FreeGameHotkeys() {
			foreach (KeyValuePair<Keys, ActionStashSpeechArgs> item in voiceHotkeys) {
				HotKeyManager.UnregisterHotKey(item.Value._unregID);
			}
			voiceHotkeys.Clear();
		}

		/// <summary>
		/// Removes all custom hotkeys
		/// </summary>
		public void FreeCustomHotkeys() {
			foreach (int key in DefinitionParser.instance.hotkeyParser.activeKeyIDs) {
				HotKeyManager.UnregisterHotKey(key);
			}
		}

		/// <summary>
		/// Removes all hotkeys except custom item ones, expensive call!
		/// </summary>
		public void FreeNonCustomHotkeys() {
			FreeControlHotkeys();
			List<Keys> toRemove = new List<Keys>();
			foreach (KeyValuePair<Keys, ActionStashSpeechArgs> item in voiceHotkeys) {
				if (!DefinitionParser.instance.hotkeyParser.currKeys.Contains(item.Key)) {
					HotKeyManager.UnregisterHotKey(item.Value._unregID);
					toRemove.Add(item.Key);
				}
			}
			for (int i = 0; i < toRemove.Count; i++) {
				voiceHotkeys.Remove(toRemove[i]);
			}
		}

		/// <summary>
		/// Selectively remove a hotkey
		/// </summary>
		public void FreeSpecific(Keys hotkey, bool unsubscribe, bool debug = false) {
			if (voiceHotkeys.ContainsKey(hotkey)) {
				if (unsubscribe) {
					HotKeyManager.UnregisterHotKey(voiceHotkeys[hotkey]._unregID);
				}

				voiceHotkeys.Remove(hotkey);
			}
			else if (controlHotkeys.ContainsKey(hotkey)) {
				if (unsubscribe) {
					HotKeyManager.UnregisterHotKey(controlHotkeys[hotkey]._unregID);
				}

				controlHotkeys.Remove(hotkey);
			}
			else {
				if (debug) { Console.WriteLine(hotkey + " not found in any list!"); }
			}
		}
		#endregion

		public void RemapHotkey(Keys key, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			FreeSpecific(key, false);
			AssignToHotkey(key, action, arguments);
		}

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

		private static void AbortReadLine(string command) {
			Program.currCommand = command;
			Thread.Sleep(250);
			PostMessage(Process.GetCurrentProcess().MainWindowHandle, 0x100, 0x0D, 0);
		}

		public void EnemyHandlingItemDroppedWrapper(SpeechRecognizedArgs args) {
			Console.WriteLine("Activated hotkey for item " + args.text + "!");
			Program.gameRecognizer.enemyHandling.ItemDropped(args.text, 1);
		}

		private void HotKeyManager_HotKeyPressed(object sender, HotKeyEventArgs e) {
			if (controlHotkeys.ContainsKey(e.Key)) {
				ActionStashString stash = controlHotkeys[e.Key];
				if (stash._keyModifier == e.Modifiers && !stash._isInactive) {
					stash._action.Invoke(stash._data);
				}
				else if (stash._isInactive) {
					Console.WriteLine("This command in currently inaccessible");
				}
			}
			else if (voiceHotkeys.ContainsKey(e.Key)) {
				ActionStashSpeechArgs stash = voiceHotkeys[e.Key];
				if (stash._keyModifier == e.Modifiers && !stash._isInactive) {
					stash._action.Invoke(stash._data);
				}
				else if (stash._isInactive) {
					Console.WriteLine("This command in currently inaccessible");
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
			public int _unregID;
			public bool _isInactive;
		}

		private struct ActionStashString {
			public KeyModifiers _keyModifier;
			public Action<string> _action;
			public string _data;
			public int _unregID;
			public bool _isInactive;
		}
	}
}
