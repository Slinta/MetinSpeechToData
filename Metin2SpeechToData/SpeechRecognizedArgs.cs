using System;

namespace Metin2SpeechToData {
	public struct SpeecRecognizedArgs {
		public SpeecRecognizedArgs(string text, float confidence) : this() {
			this.text = text;
			this.confidence = confidence;
		}

		public string text { get;}
		public float confidence { get;}
	}

	public struct PostMessageArgs {
		public PostMessageArgs(IntPtr ptr, UInt32 message, int arg1, int arg2) : this() {
			this.ptr = ptr;
			this.message = message;
			this.arg1 = arg1;
			this.arg2 = arg2;
		}

		public IntPtr ptr { get; }
		public UInt32 message { get; }
		public int arg1 { get; }
		public int arg2 { get; }

	}
}
