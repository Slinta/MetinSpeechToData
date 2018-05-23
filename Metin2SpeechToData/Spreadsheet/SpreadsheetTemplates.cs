﻿using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Metin2SpeechToData.Structures;
using System;

namespace Metin2SpeechToData {
	public class SpreadsheetTemplates : IDisposable {
		private const short columnOffset = 4;
		public enum SpreadsheetPresetType {
			MAIN,
			AREA,
			ENEMY,
			SESSION
		}

		private readonly SpreadsheetInteraction interaction;
		private ExcelPackage package;
		public SpreadsheetTemplates(SpreadsheetInteraction interaction) {
			this.interaction = interaction;
		}


		public ExcelWorksheets LoadTemplates() {
			FileInfo templatesFile = new FileInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar +
									  "Templates" + Path.DirectorySeparatorChar + "Templates.xlsx");

			if (!templatesFile.Exists) {
				throw new CustomException("Unable to locate 'Templates.xlsx' inside Templates folder. Reinstall might be necessary");
			}
			package = new ExcelPackage(templatesFile);
			return package.Workbook.Worksheets;
		}

		public Dicts InitializeMobSheet(string mobName, MobAsociatedDrops data) {

			Dicts d = new Dicts(true);

			interaction.InsertValue(new ExcelCellAddress(1, 1), "Spreadsheet for enemy: " + interaction.currentSheet.Name);
			interaction.InsertValue(new ExcelCellAddress(1, 4), "Num killed:");
			interaction.InsertValue(new ExcelCellAddress(1, 5), 0);

			Dictionary<string, string[]> itemEntries = Program.gameRecognizer.enemyHandling.mobDrops.GetDropsForMob(mobName);

			ExcelCellAddress startAddr = new ExcelCellAddress("A2");
			foreach (string key in itemEntries.Keys) {
				interaction.currentSheet.Cells[startAddr.Address].Value = key;

				for (int i = 0; i < itemEntries[key].Length; i++) {
					int _offset = i + 1;
					ExcelCellAddress itemName = new ExcelCellAddress(startAddr.Row + _offset, startAddr.Column);
					ExcelCellAddress yangVal = new ExcelCellAddress(startAddr.Row + _offset, startAddr.Column + 1);
					ExcelCellAddress totalDroped = new ExcelCellAddress(startAddr.Row + _offset, startAddr.Column + 2);
					interaction.InsertValue(itemName, itemEntries[key][i]);
					d.addresses.Add(itemEntries[key][i], totalDroped);
					interaction.InsertValue(yangVal, DefinitionParser.instance.currentGrammarFile.GetYangValue(itemEntries[key][i]));
					interaction.InsertValue(totalDroped, 0);
				}
				startAddr = new ExcelCellAddress(2, startAddr.Column + 4);
			}
			interaction.Save();
			return d;
		}

		public void InitializeMainSheet() {
			interaction.currentSheet.Cells[10, 10].Value = "TEMP INIT MAIN";
		}

		public Dicts InitializeAreaSheet(DefinitionParserData data) {
			Dicts d = new Dicts(true);

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
				SpreadsheetInteraction.Group g = new SpreadsheetInteraction.Group(address, new ExcelCellAddress(address.Row + 1, address.Column),
																				  new ExcelCellAddress(address.Row + 1, address.Column + 1),
																				  new ExcelCellAddress(address.Row + 1, address.Column + 2));
				d.groups.Add(group, g);
			}
			foreach (DefinitionParserData.Item entry in data.entries) {
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
			}
			interaction.Save();
			return d;
		}

		public (Dicts,ExcelWorksheet) InitAreaSheet(DefinitionParserData data) {
			return (default(Dicts), default(ExcelWorksheet));
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				try {
					package.Dispose();
				}
				catch {
					//Attempt to dispose
				}
				disposedValue = true;
			}
		}

		~SpreadsheetTemplates() {
			Dispose(false);
		}

		// This code added to correctly implement the disposable pattern.
		public void Dispose() {
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
