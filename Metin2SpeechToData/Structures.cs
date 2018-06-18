using OfficeOpenXml;

namespace Metin2SpeechToData.Structures {
	public struct SpeechRecognizedArgs {
		public SpeechRecognizedArgs(string text, float confidence, bool asHotkey = false) {
			this.text = text;
			this.confidence = confidence;
			this.asHotkey = asHotkey;
			this.textEnm = CCommands.GetEnum(text);
		}

		public CCommands.Speech textEnm { get; }
		public bool asHotkey { get; }
		public string text { get; }
		public float confidence { get; }
	}

	public struct ItemInsertion {
		public ItemInsertion(string itemName, int count) {
			this.itemName = itemName;
			this.count = count;
		}
		public string itemName { get; }
		public int count { get; }
	}
}
