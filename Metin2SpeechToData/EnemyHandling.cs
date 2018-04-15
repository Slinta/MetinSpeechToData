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

		}
		public void EnemyFighting(string enemy) {
			if (state == EnemyState.Fighting) {
				if(currentEnemy != enemy) {
					state = EnemyState.Fighting;
					Program.interaction.OpenWorksheet(enemy);
					Program.interaction.InitialiseWorksheet();
				}
				else {
					state = EnemyState.Fighting;
					Program.interaction.AddNumberTo(new ExcelCellAddress(1,5), 1);
				}
			}
		}

		public void EnemyFinished(string enemy) {
			state = EnemyState.NoEnemy;
			Program.interaction.AddNumberTo(new ExcelCellAddress(1, 5), 1);
		}

		public void Drop(string item) {
			Program.interaction.AddNumberTo(Program.interaction.AddressFromName(item), 1);
		}
	}
}
