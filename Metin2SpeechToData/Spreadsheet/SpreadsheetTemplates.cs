using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml;
using System;

namespace Metin2SpeechToData {
	public class SpreadsheetTemplates {
		private const short columnOffset = 4;
		public enum SpreadsheetPresetType {
			MAIN,
			AREA,
			ENEMY
		}

		private SpreadsheetInteraction interaction;

		public SpreadsheetTemplates(SpreadsheetInteraction interaction) {
			this.interaction = interaction;
		}

		public SpreadsheetHelper.Dicts InitializeMobSheet(string mobName, MobAsociatedDrops data) {

			SpreadsheetHelper.Dicts d = new SpreadsheetHelper.Dicts() {
				addresses = new Dictionary<string, ExcelCellAddress>(),
				groups = new Dictionary<string, SpreadsheetInteraction.Group>()
			};

			interaction.InsertValue(new ExcelCellAddress(1, 1), "Spreadsheet for enemy: " + interaction.currentSheet.Name);
			interaction.InsertValue(new ExcelCellAddress(1, 4), "Num killed:");
			interaction.InsertValue(new ExcelCellAddress(1, 5), 0);

			string[] itemEntries = Program.enemyHandling.mobDrops.GetDropsForMob(mobName);

			ExcelCellAddress startAddr = new ExcelCellAddress("A2");
			for (int i = 0; i < itemEntries.Length; i++) {
				ExcelCellAddress itemName = new ExcelCellAddress(startAddr.Row + i, startAddr.Column);
				ExcelCellAddress yangVal = new ExcelCellAddress(startAddr.Row + i, startAddr.Column + 1);
				ExcelCellAddress totalDroped = new ExcelCellAddress(startAddr.Row + i, startAddr.Column + 2);
				interaction.InsertValue(itemName, itemEntries[i]);
				d.addresses.Add(itemEntries[i], totalDroped);
				interaction.InsertValue(yangVal, DefinitionParser.instance.currentGrammarFile.GetYangValue(itemEntries[i]));
				interaction.InsertValue(totalDroped, 0);
			}
			interaction.Save();
			return d;
		}

		public void InitializeMainSheet() {
			interaction.currentSheet.Cells[10, 10].Value = "TEMP INIT MAIN";
		}

		public SpreadsheetHelper.Dicts InitializeAreaSheet(DefinitionParserData data) {
			SpreadsheetHelper.Dicts d = new SpreadsheetHelper.Dicts {
				addresses = new Dictionary<string, ExcelCellAddress>(),
				groups = new Dictionary<string, SpreadsheetInteraction.Group>()
			};

			interaction.InsertValue(new ExcelCellAddress(1, 1), "Spreadsheet for zone: " + interaction.currentSheet.Name);

			int[] rowOfEachGroup = new int[data.groups.Length];
			int[] columnOfEachGroup = new int[data.groups.Length];
			int groupcounter = 0;
			foreach (string group in data.groups) {
				rowOfEachGroup[groupcounter] = 2;
				columnOfEachGroup[groupcounter] = groupcounter * columnOffset + 1;
				ExcelCellAddress address = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter]);
				interaction.currentSheet.Select(new ExcelAddress(address.Row, address.Column, address.Row, address.Column + 2));
				ExcelRange r = interaction.currentSheet.SelectedRange;
				r.Merge = true;
				r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				interaction.InsertValue(address, group);
				groupcounter += 1;
				SpreadsheetInteraction.Group g = new SpreadsheetInteraction.Group {
					groupName = address,
					elementNameFirstIndex = new ExcelCellAddress(address.Row + 1, address.Column),
					yangValueFirstIndex = new ExcelCellAddress(address.Row + 1, address.Column + 1),
					totalCollectedFirstIndex = new ExcelCellAddress(address.Row + 1, address.Column + 2),
				};
				d.groups.Add(group, g);
			}
			foreach (DefinitionParserData.Entry entry in data.entries) {
				SpreadsheetInteraction.Group g = d.groups[entry.group];
				g.totalEntries++;
				d.groups[entry.group] = g;
				for (int i = 0; i < data.groups.Length; i++) {
					if (data.groups[i] == entry.group) {
						groupcounter = i;
					}
				}
				rowOfEachGroup[groupcounter] += 1;
				ExcelCellAddress nameAddr = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter]);
				ExcelCellAddress yangVal = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter] + 1);
				ExcelCellAddress collected = new ExcelCellAddress(rowOfEachGroup[groupcounter], columnOfEachGroup[groupcounter] + 2);

				interaction.InsertValue(nameAddr, entry.mainPronounciation);
				interaction.InsertValue(yangVal, entry.yangValue);
				interaction.InsertValue(collected, 0);

				interaction.currentSheet.Cells[yangVal.Address].Style.Numberformat.Format = "###,###,###,###,###";
				d.addresses.Add(entry.mainPronounciation, collected);
				interaction.Save();
			}
			return d;
		}

		public void AddItemEntry(ExcelWorksheet sheet, DefinitionParserData.Entry entry) {
			//TODO implement group sorting for mob lists
			ExcelCellAddress current = new ExcelCellAddress(2, 1 + Program.enemyHandling.mobDrops.GetGroupNumberForEnemy(sheet.Name,entry.group) * 4);
			int maxDetph = 10;
			
			while(sheet.Cells[current.Address].Value != null) {
				current = new ExcelCellAddress(current.Row + 1, current.Column);
				if(current.Row >= maxDetph) {
					current = new ExcelCellAddress(2, current.Column + 4);
				}
			}
			interaction.InsertValue(current, entry.mainPronounciation);
			interaction.InsertValue(new ExcelCellAddress(current.Row, current.Column + 1), entry.yangValue);
			interaction.InsertValue(new ExcelCellAddress(current.Row, current.Column + 2), 0);
			interaction.AddSheetToAddressEntry(sheet.Name, entry.mainPronounciation, new ExcelCellAddress(current.Row, current.Column + 2));
		}

		public void RemoveItemEntry(ExcelWorksheet sheet, DefinitionParserData.Entry entry) {
			//TODO implement group sorting for mob lists
			ExcelCellAddress current = new ExcelCellAddress("A2");
			int maxDetph = 10;

			while (sheet.Cells[current.Address].GetValue<string>() != entry.mainPronounciation) {
				current = new ExcelCellAddress(current.Row + 1, current.Column);
				if (current.Row >= maxDetph) {
					current = new ExcelCellAddress(2, current.Column + 4);
				}
			}
			//Found the cell
			interaction.InsertValue<object>(current, null);
			interaction.InsertValue<object>(new ExcelCellAddress(current.Row, current.Column + 1), null);
			interaction.InsertValue<object>(new ExcelCellAddress(current.Row, current.Column + 2), null);
			interaction.RemoveSheetToAddressEntry(sheet.Name, entry.mainPronounciation);
			//TODO: Shof the cell upwards to remove gaps
		}
	}
}
