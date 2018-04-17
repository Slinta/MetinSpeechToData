using System;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public class EnemyHandling {
		public enum EnemyState {
			NoEnemy,
			Fighting
		}

		public EnemyState state;
		private string currentEnemy = "";

		public EnemyHandling() {
			if (Program.debug) {
				WrittenControl.OnModifierWordHear += EnemyTargetingModifierRecognized;
			}
			Program.OnModifierWordHear += EnemyTargetingModifierRecognized;
		}

		~EnemyHandling() {
			if (Program.debug) {
				WrittenControl.OnModifierWordHear += EnemyTargetingModifierRecognized;
			}
			Program.OnModifierWordHear -= EnemyTargetingModifierRecognized;
		}

		/// <summary>
		/// Event fired after a modifier word "New Target" was said
		/// </summary>
		/// <param name="keyWord">In this case always equal to "NEW_TARGET"</param>
		/// <param name="args">In this case always contains the enemy name/ambiguity at [0]</param>
		public void EnemyTargetingModifierRecognized(SpeechRecognitionHelper.ModifierWords keyWord, params string[] args) {
			switch (state) {
				case EnemyState.NoEnemy: {
					//Initial state
					if(args[0] == "" && currentEnemy == "") {
						return;
					}
					string actualEnemyName = DefinitionParser.instance.currentMobGrammarFile.GetMainPronounciation(args[0]);
					state = EnemyState.Fighting;
					Program.interaction.OpenWorksheet(actualEnemyName);
					Program.interaction.InitialiseWorksheet();
					currentEnemy = actualEnemyName;
					Console.WriteLine("Acquired target: " + currentEnemy);
					break;
				}
				case EnemyState.Fighting: {
					state = EnemyState.NoEnemy;
					Console.WriteLine("Killed " + currentEnemy + ", the death noted in " + Program.interaction.currentSheet.Name);
					Program.interaction.AddNumberTo(new ExcelCellAddress(1, 5), 1);
					currentEnemy = "";
					break;
				}
			}
		}

		/// <summary>
		/// Increases number count to 'item' in current speadsheet
		/// </summary>
		public void ItemDropped(string item) {
			Program.interaction.AddNumberTo(Program.interaction.AddressFromName(item), 1);
		}
	}
}
