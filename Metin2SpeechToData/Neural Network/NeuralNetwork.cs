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

		public float learningRate = 0.1f;

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

		public void Train(double[] data, double[] expected) {
			Matrix inputs = Matrix.FromAray(data);
			Matrix hidden = input_hidden_weights * inputs;
			hidden += hidden_bias;
			hidden.Map(Sigmoid);


			Matrix outputs = hidden_output_weights * hidden;
			outputs += output_bias;
			outputs.Map(Sigmoid);
			//^^^ This would be the guess

			Matrix targets = Matrix.FromAray(expected);
			Matrix outputError = targets - outputs;
			//^^^This is the error
			Matrix.Print(outputError);

			Matrix gradient = Matrix.Map(outputs, DeriveSigmoid);
			gradient *= outputError;
			gradient *= learningRate;


			Matrix hiddenT = Matrix.Transpose(hidden);
			Matrix hidden_output_e_delta = gradient * hiddenT;

			hidden_output_weights += hidden_output_e_delta;
			output_bias += gradient;


			Matrix hidden_output_T = Matrix.Transpose(hidden_output_weights);
			Matrix hidden_error = hidden_output_T * outputError;

			Matrix hiddenGradient = Matrix.Map(hidden, DeriveSigmoid);
			hiddenGradient *= hidden_error;
			hiddenGradient *= learningRate;


			Matrix inputMatrixT = Matrix.Transpose(inputs);
			Matrix input_hidden_e_delta = hiddenGradient * inputMatrixT;

			input_hidden_weights += input_hidden_e_delta;
			hidden_bias += hiddenGradient;
		}


		private double Sigmoid(double x) {
			return 1 / (1 + Math.Exp(-x));
		}

		private double DeriveSigmoid(double x) {
			return x - (1 - x);
		}

		public struct NData {
			public double[] input;
			public double[] expected;
		}
	}
}
