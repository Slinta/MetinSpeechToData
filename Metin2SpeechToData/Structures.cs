using OfficeOpenXml;

namespace Metin2SpeechToData.Structures{
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
		public ItemInsertion(ExcelCellAddress address, int count) {
			this.address = address;
			this.count = count;
		}
		public ExcelCellAddress address { get; }
		public int count { get; }
	}
}
