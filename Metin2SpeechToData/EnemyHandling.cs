using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Metin2SpeechToData {
	public class EnemyHandling {
		public enum EnemyState {
			Fighting,
			NoEnemy,
		}
		public EnemyState state;
		private string currentEnemy;
		public EnemyHandling() {
			Program.OnModifierWordHear += EnemyFighting;
		}
		public void EnemyFighting(string enemy, params string[] args) {
			string actualEnemyName;
			try {
				actualEnemyName = args[0];
			}
			catch {
				throw new Exception("args was empty");
			}
			actualEnemyName = DefinitionParser.instance.currentMobGrammarFile.GetMainPronounciation(actualEnemyName);

				if (string.IsNullOrEmpty(currentEnemy) || currentEnemy != actualEnemyName) {
					state = EnemyState.Fighting;
					
					try{
						Program.interaction.OpenWorksheet(actualEnemyName);
						currentEnemy = actualEnemyName;
					}
					catch {
						throw new Exception("args was empty");
					}
					
					Program.interaction.InitialiseWorksheet();
				}
				else {
					throw new Exception("Enemy not finished");
				}
			
		}

		public void EnemyFinished() {
			state = EnemyState.NoEnemy;
			Console.WriteLine(currentEnemy + " killed");
			Console.WriteLine("The death noted in " + Program.interaction.xlsSheet.Name);
			Program.interaction.AddNumberTo(new ExcelCellAddress(1, 5), 1);
		}

		public void Drop(string item) {
			Program.interaction.AddNumberTo(Program.interaction.AddressFromName(item), 1);
		}
	}
}
