using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Metin2SpeechToData.Neural_Network {
	class NeuralNetwork {

		private Node[] inputNodes;

		/// <summary>
		/// Creates a new instance of a neural network
		/// </summary>
		/// <param name="inputs"> number of input nodes</param>
		/// <param name="hiddenNodes">number of hidden nodes</param>
		/// <param name="outputs">number of classifications(number of things to recognize)</param>
		public NeuralNetwork(int inputs, int hiddenNodes, int outputs) {
			inputNodes = new Node[inputs];
			for (int i = 0; i < inputs; i++) {

			}
		}

		public void Classify(double[] data) {
			if(data.Length != inputNodes.Length) {
				throw new Exception("Passing in invalid amount of data, this network accepts " + inputNodes.Length + ", sent " + data.Length);
			}
			//Classify
		}
	}

	class Node {
		public Node() {
			//Something
		}
	}
}
