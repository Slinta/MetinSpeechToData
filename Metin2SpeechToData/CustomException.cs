using System;

namespace Metin2SpeechToData {
	[Serializable]
	public class CustomException : Exception {
		public CustomException(string message) : base(message) { }

		public CustomException(string message, bool fatal) : base(message) { }

		public CustomException(string message, Exception inner) : base(message, inner) { }
	}
}
