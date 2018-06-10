using System;

namespace Metin2SpeechToData {
	public static class Quicksort {
		public static void Sort<T>(this T[] array, int start, int end, Func<T,T,int> comparer) {
			if (start < end) {
				int pivot = Partition(array, start, end, comparer);
				Sort(array, start, pivot - 1, comparer);
				Sort(array, pivot + 1, end, comparer);
			}
			
		}

		private static int Partition<T>(T[] array, int index1, int index2, Func<T, T, int> comparer) {
			int i = index1 - 1;
			for (int j = index1; j < index2; j++) {
				if(comparer(array[j],array[index2]) <= 0) {
					i++;
					Swap(array, i, j);
				}
			}
			Swap(array, ++i, index2);
			return i;
		}

		private static void Swap<T>(T[] array, int st, int nd) {
			T temp = array[st];
			array[st] = array[nd];
			array[nd] = temp;
		}
	}
}
