using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Metin2SpeechToData.Neural_Network {
	class NeuralNetwork {

		private Matrix input_hidden_weights;
		private Matrix hidden_output_weights;

		private Matrix hidden_bias;
		private Matrix output_bias;

		/// <summary>
		/// Creates a new instance of a neural network
		/// </summary>
		/// <param name="inputs"> number of input nodes</param>
		/// <param name="hiddenNodes">number of hidden nodes</param>
		/// <param name="outputs">number of classifications(number of things to recognize)</param>
		public NeuralNetwork(int inputs, int hiddenNodes, int outputs) {
			input_hidden_weights = new Matrix(hiddenNodes, inputs);
			hidden_output_weights = new Matrix(outputs, hiddenNodes);
			input_hidden_weights.InitRandomOneNormalized();
			hidden_output_weights.InitRandomOneNormalized();

			hidden_bias = new Matrix(hiddenNodes, 1);
			output_bias = new Matrix(outputs, 1);
			hidden_bias.InitRandomOneNormalized();
			output_bias.InitRandomOneNormalized();
		}

		public double[] Classify(double[] data) {

			Matrix dd = Matrix.FromAray(data);
			Matrix fromHidden = input_hidden_weights * dd;
			fromHidden += hidden_bias;
			fromHidden.Map(Sigmoid);

			Matrix _out = hidden_output_weights * fromHidden;
			_out += output_bias;
			_out.Map(Sigmoid);

			return _out.ToArray();
		}

		private double Sigmoid(double x) {
			return 1 / (1 + Math.Exp(-x));
		}

	}
}
