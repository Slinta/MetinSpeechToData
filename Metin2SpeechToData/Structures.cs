using System.Collections.Generic;
using OfficeOpenXml;

namespace Metin2SpeechToData.Structures{
	public struct SpeechRecognizedArgs {
		public SpeechRecognizedArgs(string text, float confidence, bool asHotkey = false) : this() {
			this.text = text;
			this.confidence = confidence;
			this.asHotkey = asHotkey;
		}

		public bool asHotkey { get; }
		public string text { get; }
		public float confidence { get; }
	}

	public struct Dicts {
		public Dicts(bool initialize) {
			addresses = new Dictionary<string, ExcelCellAddress>();
			groups = new Dictionary<string, SpreadsheetInteraction.Group>();
		}

		public Dictionary<string, ExcelCellAddress> addresses { get; }
		public Dictionary<string, SpreadsheetInteraction.Group> groups { get; }
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
