using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using System.Collections.Generic;
using Metin2SpeechToData.Structures;

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

	public class HotKeyManager {
		private delegate void RegisterHotKeyDelegate(IntPtr hwnd, int id, uint modifiers, uint key);
		private delegate void UnRegisterHotKeyDelegate(IntPtr hwnd, int id);

		public event EventHandler<HotKeyEventArgs> HotKeyPressed;

		private MessageWindow _wnd;
		private IntPtr _hwnd;
		private static readonly ManualResetEvent _windowReadyEvent = new ManualResetEvent(false);

		private int _id = 0;

		internal HotKeyManager() {
			Thread messageLoop = new Thread(delegate () {
				Application.Run(new MessageWindow(this, ref _wnd, ref _hwnd));
			});
			messageLoop.Name = "MessageLoopThread";
			messageLoop.IsBackground = true;
			messageLoop.Start();
		}


		public int RegisterHotKey(Keys key, KeyModifiers modifiers) {
			_windowReadyEvent.WaitOne();
			int id = Interlocked.Increment(ref _id);
			_wnd.Invoke(new RegisterHotKeyDelegate(RegisterHotKeyInternal), _hwnd, id, (uint)modifiers, (uint)key);
			return id;
		}

		public void UnregisterHotKey(int id) {
			_wnd.Invoke(new UnRegisterHotKeyDelegate(UnRegisterHotKeyInternal), _hwnd, id);
		}


		private void RegisterHotKeyInternal(IntPtr windowHandle, int id, uint modifiers, uint key) {
			NativeMethods.RegisterHotKey(windowHandle, id, modifiers, key);
		}

		private void UnRegisterHotKeyInternal(IntPtr hwnd, int id) {
			NativeMethods.UnregisterHotKey(_hwnd, id);
		}

		private void OnHotKeyPressed(HotKeyEventArgs e) {
			HotKeyPressed?.Invoke(this, e);
		}


		private class MessageWindow : Form {
			private const int WM_HOTKEY = 0x312;
			private readonly HotKeyManager _manager;

			public MessageWindow(HotKeyManager manager, ref MessageWindow _wnd, ref IntPtr _hwnd) {
				_manager = manager;
				_wnd = this;
				_hwnd = this.Handle;
				_windowReadyEvent.Set();
			}

			protected override void WndProc(ref Message m) {
				if (m.Msg == WM_HOTKEY) {
					HotKeyEventArgs e = new HotKeyEventArgs(m.LParam);
					_manager.OnHotKeyPressed(e);
				}
				base.WndProc(ref m);
			}

			protected override void SetVisibleCore(bool value) {
				// Ensure the window never becomes visible
				base.SetVisibleCore(false);
			}
		}

	}
	public static class NativeMethods {
		[DllImport("user32", SetLastError = true)]
		internal static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

		[DllImport("user32", SetLastError = true)]
		internal static extern bool UnregisterHotKey(IntPtr hWnd, int id);

		[DllImport("User32.Dll", EntryPoint = "PostMessageA")]
		internal static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, int lParam);
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

		private readonly HotKeyManager manager;

		private readonly Dictionary<Keys, ActionStashSpeechArgs> voiceHotkeys = new Dictionary<Keys, ActionStashSpeechArgs>();

		private readonly Dictionary<Keys, ActionStashString> controlHotkeys = new Dictionary<Keys, ActionStashString>();

		private readonly Dictionary<Keys, ActionData<int>> customHotkeys = new Dictionary<Keys, ActionData<int>>();

		private readonly Dictionary<Keys, ActionStashSpeechArgs> itemHotkeys = new Dictionary<Keys, ActionStashSpeechArgs>();


		public bool hotkeyOverriding { get; set; }

		public HotKeyMapper() {
			manager = new HotKeyManager();
			manager.HotKeyPressed += new EventHandler<HotKeyEventArgs>(HotKeyManager_HotKeyPressed);
		}

		#region Add Hotkeys

		/// <summary>
		/// Assign hotkey 'selectedKey' to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, int selection, Action<int> action) {
			if (customHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + customHotkeys[selectedKey] + "!");
			}
			if (customHotkeys.TryGetValue(selectedKey, out ActionData<int> data)) {
				customHotkeys.Remove(selectedKey);
			}
			customHotkeys.Add(selectedKey, new ActionData<int>() {
				action = action,
				_unregID = manager.RegisterHotKey(selectedKey, KeyModifiers.None),
				_keyModifiers = KeyModifiers.None,
				data = selection
			});
			return customHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign item hotkey
		/// </summary>
		public int AssignItemHotkey(Keys selectedKey, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (itemHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + customHotkeys[selectedKey] + "!");
			}
			if (itemHotkeys.TryGetValue(selectedKey, out ActionStashSpeechArgs data)) {
				itemHotkeys.Remove(selectedKey);
			}
			itemHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_keyModifier = KeyModifiers.None,
				_action = action,
				_unregID = manager.RegisterHotKey(selectedKey, KeyModifiers.None),
				_data = arguments,
				_isInactive = true
			});
			return itemHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign item hotkey
		/// </summary>
		public int AssignItemHotkey(Keys selectedKey, KeyModifiers modifier1, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (itemHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + customHotkeys[selectedKey] + "!");
			}
			if (itemHotkeys.TryGetValue(selectedKey, out ActionStashSpeechArgs data)) {
				itemHotkeys.Remove(selectedKey);
			}
			itemHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_keyModifier = modifier1,
				_action = action,
				_unregID = manager.RegisterHotKey(selectedKey, KeyModifiers.None),
				_data = arguments,
				_isInactive = true
			});
			return itemHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign item hotkey
		/// </summary>
		public int AssignItemHotkey(Keys selectedKey, KeyModifiers modifier1, KeyModifiers modifier2, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (itemHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + customHotkeys[selectedKey] + "!");
			}
			if (itemHotkeys.TryGetValue(selectedKey, out ActionStashSpeechArgs data)) {
				itemHotkeys.Remove(selectedKey);
			}
			itemHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_keyModifier = modifier1 | modifier2,
				_action = action,
				_unregID = manager.RegisterHotKey(selectedKey, KeyModifiers.None),
				_data = arguments,
				_isInactive = true
			});
			return itemHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign item hotkey
		/// </summary>
		public int AssignItemHotkey(Keys selectedKey, KeyModifiers modifier1, KeyModifiers modifier2, KeyModifiers modifier3, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (itemHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + customHotkeys[selectedKey] + "!");
			}
			if (itemHotkeys.TryGetValue(selectedKey, out ActionStashSpeechArgs data)) {
				itemHotkeys.Remove(selectedKey);
			}
			itemHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_keyModifier = modifier1 | modifier2 | modifier3,
				_action = action,
				_unregID = manager.RegisterHotKey(selectedKey, KeyModifiers.None),
				_data = arguments,
				_isInactive = true
			});
			return itemHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			if (voiceHotkeys.TryGetValue(selectedKey, out ActionStashSpeechArgs data)) {
				voiceHotkeys.Remove(selectedKey);
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = KeyModifiers.None,
				_unregID = manager.RegisterHotKey(selectedKey, KeyModifiers.None),
				_isInactive = false
			}
			);
			return voiceHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' + a 'modifier' key to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, KeyModifiers modifier, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			if (voiceHotkeys.TryGetValue(selectedKey, out ActionStashSpeechArgs data)) {
				voiceHotkeys.Remove(selectedKey);
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = modifier,
				_unregID = manager.RegisterHotKey(selectedKey, modifier)
			}
			);
			return voiceHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' + 'modifier' keys to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, KeyModifiers modifier1, KeyModifiers modifier2, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			if (voiceHotkeys.TryGetValue(selectedKey, out ActionStashSpeechArgs data)) {
				voiceHotkeys.Remove(selectedKey);
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = modifier1 | modifier2,
				_unregID = manager.RegisterHotKey(selectedKey, modifier1 | modifier2)
			}
			);
			return voiceHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' + 'modifier' keys to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, KeyModifiers modifier1, KeyModifiers modifier2, KeyModifiers modifier3, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			if (voiceHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + voiceHotkeys[selectedKey] + "!");
			}
			if (voiceHotkeys.TryGetValue(selectedKey, out ActionStashSpeechArgs data)) {
				voiceHotkeys.Remove(selectedKey);
			}
			voiceHotkeys.Add(selectedKey, new ActionStashSpeechArgs() {
				_action = action,
				_data = arguments,
				_keyModifier = modifier1 | modifier2 | modifier3,
				_unregID = manager.RegisterHotKey(selectedKey, modifier1 | modifier2 | modifier3)
			}
			);
			return voiceHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, string command) {
			if (controlHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + controlHotkeys[selectedKey] + "!");
			}
			if (controlHotkeys.TryGetValue(selectedKey, out ActionStashString data)) {
				voiceHotkeys.Remove(selectedKey);
			}
			controlHotkeys.Add(selectedKey, new ActionStashString() {
				_action = AbortReadLineAndCallCommand,
				_data = command,
				_keyModifier = KeyModifiers.None,
				_unregID = manager.RegisterHotKey(selectedKey, KeyModifiers.None)
			}
			);
			return controlHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' + a 'modifier' key to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, KeyModifiers modifier, string command) {
			if (controlHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + controlHotkeys[selectedKey] + "!");
			}
			if (controlHotkeys.TryGetValue(selectedKey, out ActionStashString data)) {
				controlHotkeys.Remove(selectedKey);
			}
			controlHotkeys.Add(selectedKey, new ActionStashString() {
				_action = AbortReadLineAndCallCommand,
				_data = command,
				_keyModifier = modifier,
				_unregID = manager.RegisterHotKey(selectedKey, modifier)
			}
			);
			return controlHotkeys[selectedKey]._unregID;
		}

		/// <summary>
		/// Assign hotkey 'selectedKey' + 'modifier' keys to call function 'action' with 'arguments'
		/// </summary>
		public int AssignToHotkey(Keys selectedKey, KeyModifiers modifier1, KeyModifiers modifier2, string command) {
			if (controlHotkeys.ContainsKey(selectedKey) && !hotkeyOverriding) {
				throw new CustomException(selectedKey + " already mapped to " + controlHotkeys[selectedKey] + "!");
			}
			if (controlHotkeys.TryGetValue(selectedKey, out ActionStashString data)) {
				controlHotkeys.Remove(selectedKey);
			}
			controlHotkeys.Add(selectedKey, new ActionStashString() {
				_action = AbortReadLineAndCallCommand,
				_data = command,
				_keyModifier = modifier1 | modifier2,
				_unregID = manager.RegisterHotKey(selectedKey, modifier1 | modifier2)
			}
			);
			return controlHotkeys[selectedKey]._unregID;
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
				manager.UnregisterHotKey(item.Value._unregID);
			}
			controlHotkeys.Clear();
		}

		/// <summary>
		/// Removes all game hotkeys
		/// </summary>
		public void FreeGameHotkeys() {
			foreach (KeyValuePair<Keys, ActionStashSpeechArgs> item in voiceHotkeys) {
				manager.UnregisterHotKey(item.Value._unregID);
			}
			voiceHotkeys.Clear();
		}

		/// <summary>
		/// Removes all item hotkeys
		/// </summary>
		public void FreeItemHotkeys() {
			foreach (KeyValuePair<Keys, ActionStashSpeechArgs> item in itemHotkeys) {
				manager.UnregisterHotKey(item.Value._unregID);
			}
			itemHotkeys.Clear();
		}

		/// <summary>
		/// Toggles Item hotkeys
		/// </summary>
		/// <param name="state"></param>
		public void ToggleItemHotkeys(bool state) {
			if (!state) {
				List<Keys> coll = new List<Keys>();

				foreach (KeyValuePair<Keys, ActionStashSpeechArgs> item in itemHotkeys) {
					manager.UnregisterHotKey(item.Value._unregID);
					coll.Add(item.Key);
				}
				foreach (Keys key in coll) {
					SetInactive(key, true);
				}
			}
			else {
				List<Keys> coll = new List<Keys>();
				foreach (KeyValuePair<Keys, ActionStashSpeechArgs> item in itemHotkeys) {
					manager.RegisterHotKey(item.Key, item.Value._keyModifier);
					coll.Add(item.Key);
				}
				foreach (Keys key in coll) {
					SetInactive(key, false);
				}
			}
		}

		/// <summary>
		/// Selectively remove a hotkey
		/// </summary>
		public void FreeSpecific(Keys hotkey, bool unsubscribe, bool debug = false) {
			if (voiceHotkeys.ContainsKey(hotkey)) {
				if (unsubscribe) {
					manager.UnregisterHotKey(voiceHotkeys[hotkey]._unregID);
				}
				voiceHotkeys.Remove(hotkey);
			}
			else if (controlHotkeys.ContainsKey(hotkey)) {
				if (unsubscribe) {
					manager.UnregisterHotKey(controlHotkeys[hotkey]._unregID);
				}
				controlHotkeys.Remove(hotkey);
			}
			else if (customHotkeys.ContainsKey(hotkey)) {
				if (unsubscribe) {
					manager.UnregisterHotKey(customHotkeys[hotkey]._unregID);
				}
				customHotkeys.Remove(hotkey);
			}
			else {
				if (debug) { Console.WriteLine(hotkey + " not found in any list!"); }
			}
		}

		/// <summary>
		/// Unregisters hotkey by given ID
		/// </summary>
		public void FreeSpecific(int unregID) {
			manager.UnregisterHotKey(unregID);
		}

		#endregion

		public void RemapHotkey(Keys key, Action<SpeechRecognizedArgs> action, SpeechRecognizedArgs arguments) {
			FreeSpecific(key, true);
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
			if (itemHotkeys.ContainsKey(key)) {
				ActionStashSpeechArgs stash = itemHotkeys[key];
				stash._isInactive = state;
				itemHotkeys[key] = stash;
			}
		}

		public static void AbortReadLine() {
			NativeMethods.PostMessage(Process.GetCurrentProcess().MainWindowHandle, 0x100, 0x0D, 0);
		}

		#region Item Hotkey Handling

		public static void AbortReadLineAndCallCommand(string command) {
			Program.currCommand = command;
			//Thread.Sleep(250);
			AbortReadLine();
		}

		private GameRecognizer recognizer;
		public void AttachHotkeyWrapper(GameRecognizer rec) {
			recognizer = rec;
		}
		public void DetachHotkeyWrapper() {
			recognizer = null;
		}
		public void EnemyHandlingItemDroppedWrapper(SpeechRecognizedArgs args) {
			Console.WriteLine("Activated hotkey for item " + args.text + "!");
			recognizer.enemyHandling.ItemDropped(args.text, 1);
		}
		#endregion

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
			else if (customHotkeys.ContainsKey(e.Key)) {
				ActionData<int> data = customHotkeys[e.Key];
				data.action.Invoke(data.data);
			}
			else if (itemHotkeys.ContainsKey(e.Key)) {
				ActionStashSpeechArgs stash = itemHotkeys[e.Key];
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

		private struct ActionData<T> {
			public KeyModifiers _keyModifiers;
			public Action<T> action;
			public T data;
			public int _unregID;
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
