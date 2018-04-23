namespace Metin2SpeechToData {
	class SpreadsheetConstants {

		public const string DEFAULT_SHEET = "Metin2 Drop Analyzer";

		public static int DigitCount(int i) {
			int count = 1;
			for (int j = 0; j < int.MaxValue; j++) {
				int newVal = i / 10;
				if (newVal >= 1) {
					count++;
					i = newVal;
				}
				else {
					break;
				}
			}
			return count;
		}

		public static double GetCellWidth(int number, bool addCurrencyOffset) {
			int count = DigitCount(number);
			int spaces = count / 3;
			double width = addCurrencyOffset ? 4 : 2;
			width += spaces + count;
			return width;
		}
	}
}
