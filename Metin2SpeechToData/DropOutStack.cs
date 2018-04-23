namespace Metin2SpeechToData {
	/* @ Greg Dean
	 * https://stackoverflow.com/questions/384042/can-i-limit-the-depth-of-a-generic-stack
	 */

	class DropOutStack<T> {
		private T[] items;
		private int top = 0;

		/// <summary>
		/// Creates a stack with 'capacity' depth, items added beyond this capacity will overwrite oldest entry
		/// </summary>
		public DropOutStack(int capacity) {
			items = new T[capacity];
		}

		/// <summary>
		/// Push item onto the stack, possibly overwrite oldest if capacity is exceeded
		/// </summary>
		public void Push(T item) {
			items[top] = item;
			top = (top + 1) % items.Length;
		}

		/// <summary>
		/// Pop item off the stack, ater depletion return defaults of 'T'
		/// </summary>
		public T Pop() {
			top = (items.Length + top - 1) % items.Length;
			T output = items[top];
			items[top] = default(T);
			return output;
		}
	}
}
