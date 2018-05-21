using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sheet_DefinitionValueSync {
	public static class WordSimilarity {
		/// <summary>
		/// Compute the distance between two strings.
		/// </summary>
		public static int Compute(string fist, string second) {
			int st_length = fist.Length;
			int nd_length = second.Length;
			int[,] distance_matrix = new int[st_length + 1, nd_length + 1];
			// Step 1
			if (st_length == 0) {
				return nd_length;
			}

			if (nd_length == 0) {
				return st_length;
			}

			for (int i = 0; i <= st_length; distance_matrix[i, 0] = i++) { /*Populate first column*/ }
			for (int j = 0; j <= nd_length; distance_matrix[0, j] = j++) { /*Populate first row*/ }

			// Step 3
			for (int i = 1; i <= st_length; i++) {
				//Step 4
				for (int j = 1; j <= nd_length; j++) {
					// Step 5
					int cost = (second[j - 1] == fist[i - 1]) ? 0 : 1;
					distance_matrix[i, j] = Math.Min(
						Math.Min(distance_matrix[i - 1, j] + 1, distance_matrix[i, j - 1] + 1),
						distance_matrix[i - 1, j - 1] + cost);
				}
			}
			return distance_matrix[st_length, nd_length];
		}
	}
}
