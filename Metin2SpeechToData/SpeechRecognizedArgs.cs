namespace Metin2SpeechToData {
	public struct SpeechRecognizedArgs {
		public SpeechRecognizedArgs(string text, float confidence) : this() {
			this.text = text;
			this.confidence = confidence;
		}

		public string text { get; }
		public float confidence { get; }
	}
}
