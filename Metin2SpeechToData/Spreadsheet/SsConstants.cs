﻿namespace Metin2SpeechToData.Spreadsheet {
	public static class SsConstants {
		public static class SsControl {
			/// <summary>
			/// Name of the sheet, works for all sheet types
			/// </summary>
			public const string C_SHEET_NAME = "B2";
			/// <summary>
			/// Return link to main sheet, works for all sheet types
			/// </summary>
			public const string C_RETURN_LINK = "B5";
			/// <summary>
			/// Last session that was merged to this sheet, works for AREA and ENEMY
			/// </summary>
			public const string C_LAST_SESSION_LINK = "B7";

			/// <summary>
			/// Sum of all yangs gaind from this sheet, works for all sheet types
			/// </summary>
			public const string C_TOTAL_DROPED_VALUE = "P5";
			/// <summary>
			/// Number of groups, works for AREA and ENEMY
			/// </summary>
			public const string A_E_TOTAL_GROUPS = "P6";
			/// <summary>
			/// Total item count on this sheet (droped)
			/// </summary>
			public const string C_TOTAL_ITEMS = "P7";
			/// <summary>
			/// Total merged sesstions to this sheet, works for AREA and ENEMY
			/// </summary>
			public const string A_E_TOTAL_MERGED_SESSIONS = "P8";
			/// <summary>
			/// Last modification to this sheet
			/// </summary>
			public const string A_E_LAST_MODIFICATION = "P9";
			/// <summary>
			/// Link to asociated enemy sheet, works only for AREA
			/// </summary>
			public const string A_ENEMIES_FIRST_LINK = "F5";


			public const string AREA_TOTAL_DROPED_ITEMS = "P4";
		}

		public static class SsSession {
			public const string C_SHEET_NAME = "B2";
			public const string C_RETURN_LINK = "B5";
		}

		public const int GROUP_COL = 3;
		public const int GROUP_ROW = 12;

		public const int ITEM_COL = 3;
		public const int ITEM_ROW = 13;

		public const string UNSPEICIFIED_ENEMY = "Not Specified";

		public const string DATA_FIRST_ENTRY = "C13";

		public const string MAIN_HLINK_OFFSET = "A1";
		public const string MAIN_FIRST_HLINK = "B2";

		public const byte H_FIRST_ROW = 3/*13*/;
		public const byte H_COLUMN_INCREMENT = 4;

		public const string H_DEFAULT_SHEET_NAME = "Metin2 Drop Analyzer";
	}
}
