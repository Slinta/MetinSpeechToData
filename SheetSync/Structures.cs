using OfficeOpenXml;

namespace SheetSync {
	class Structures {
		internal struct Typos {
			public Typos(string word, string[] possible_replacements, ExcelCellAddress sheetLocation, ExcelWorksheet sheet) {
				originalTypo = word;
				alternatives = possible_replacements;
				location = sheetLocation;
				this.sheet = sheet;
			}
			public string originalTypo { get; }
			public string[] alternatives { get; }
			public ExcelCellAddress location { get; }
			public ExcelWorksheet sheet { get; }

		}


		internal struct Item {
			public Item(string name, string fileOrigin, int fileLine, uint yangValue) {
				this.name = name;
				this.fileOrigin = fileOrigin;
				this.fileLine = fileLine;
				this.yangValue = yangValue;
			}

			public string name { get; }
			public string fileOrigin { get; }
			public int fileLine { get; }
			public uint yangValue { get; }
		}

		internal struct Diffs {
			public Diffs(ExcelWorksheet sheet, ExcelCellAddress address, string fileOrigin, int fileLine, uint sheetVal, uint fileVal) {
				currentSheet = sheet;
				location = address;
				itemDef = fileLine;
				itemFile = fileOrigin;
				sheetYangVal = sheetVal;
				fileYangVal = fileVal;
			}

			public string itemFile { get; }
			public ExcelWorksheet currentSheet { get; }
			public ExcelCellAddress location { get; }
			public int itemDef { get; }
			public uint sheetYangVal { get; }
			public uint fileYangVal { get; }
		}
	}
}
