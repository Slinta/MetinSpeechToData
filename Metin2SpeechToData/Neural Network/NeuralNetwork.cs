using System;
using System.Collections.Generic;

namespace Metin2SpeechToData.Neural_Network {
	class NeuralNetwork {

		public const float learningRate = 0.1f;

		Layer inputLayer;
		List<Layer> hidden_Layers;
		Layer outputLayer;

		/// <summary>
		/// Create a new neural network
		/// </summary>
		/// <param name="inputLength">NO. of input neurons</param>
		/// <param name="hiddenLayersCount">NO. of hidden layers</param>
		/// <param name="hiddenLayersLength">NO. of neurons in each hidden layer</param>
		/// <param name="outputLayersLength">NO. of output neurons</param>
		public NeuralNetwork(int inputLength, int hiddenLayersCount, int hiddenLayersLength, int outputLayersLength) {
			inputLayer = new Layer(inputLength, 0);
			hidden_Layers = new List<Layer>();
			for (int i = 1; i < hiddenLayersCount + 1; i++) {
				hidden_Layers.Add(new Layer(hiddenLayersLength, i));
			}
			outputLayer = new Layer(outputLayersLength, hiddenLayersCount + 1);
			MakeConnections(inputLayer, hidden_Layers, outputLayer);
		}


		/// <summary>
		/// Connects all layers
		/// </summary>
		private void MakeConnections(Layer input, List<Layer> hidden, Layer output) {
			Random r = new Random(Environment.TickCount);
			input.ConnectLayer(hidden[0], r.Next(1000));
			for (int i = 1; i < hidden.Count - 1; i++) {
				hidden[i].ConnectLayer(hidden[i + 1], r.Next(1000));
			}
			hidden[hidden.Count - 1].ConnectLayer(output, r.Next(1000));
		}

		/// <summary>
		/// Processes input and returns a guess
		/// </summary>
		/// <param name="data">double array of size 'inputNeuronsCount'</param>
		public double[] FeedForward(double[] data) {
			return inputLayer.ProcessInput(Matrix.FromAray(data)).ToArray();
		}


		public void Train(double[] input, double[] expected, bool debug = false) {
			Matrix expectedMatrix = Matrix.FromAray(expected);
			Matrix outputMatrix = Matrix.FromAray(FeedForward(input));
			Matrix finalError = expectedMatrix - outputMatrix; // distribute this to lasthidden/output connection

			if (debug) {
				Console.WriteLine("Expecting:");
				Matrix.Print(expectedMatrix);
				Console.WriteLine("Instead got:");
				Matrix.Print(outputMatrix);
				Console.WriteLine("That is off by...");
				Matrix.Print(finalError);
			}

			Matrix weightMatrixToOutputTransposed = Matrix.Transpose(outputLayer.inputConnection.getConnectionMatrix);

			if (debug) {
				Console.WriteLine("Transpose connection matrix inputs to the output layer");
				Matrix.Print(weightMatrixToOutputTransposed);
			}

			Matrix lastHiddenLayerError = weightMatrixToOutputTransposed * finalError;

			if (debug) {
				Console.WriteLine("The error that goes to connection from last hidden to output");
				Matrix.Print(lastHiddenLayerError);
				Console.WriteLine("Gradient descent");
			}

			Matrix toOutputGradient = Matrix.Map(outputMatrix, DeriveSigmoid);
			toOutputGradient *= finalError;
			toOutputGradient *= learningRate;

			Matrix lastStepWeightMatrix = Matrix.Transpose(hidden_Layers[hidden_Layers.Count -1].inputConnection.getConnectionMatrix);

			if (debug) {
				Console.WriteLine("What is stored in the last hidden layer's storage TRANSPOSED");
				Matrix.Print(lastStepWeightMatrix);
			}

			Matrix lastErrorDelta = toOutputGradient * lastStepWeightMatrix;

			if (debug) {
				Console.WriteLine("Apply gradient to our transposed matrix");
				Matrix.Print(lastErrorDelta);
				Console.WriteLine("Adjusted connection from:");
				Matrix.Print(outputLayer.inputConnection.getConnectionMatrix);
			}

			outputLayer.inputConnection.AdjustByDelta(lastErrorDelta);
			outputLayer.inputConnection.AdjustBias(toOutputGradient);

			if (debug) {
				Console.WriteLine("To:");
				Matrix.Print(outputLayer.inputConnection.getConnectionMatrix);
			}

			Matrix currentLayerError = lastHiddenLayerError;

			if (debug) {
				Console.WriteLine("The error that goes deeper:");
				Matrix.Print(currentLayerError);
			}

			for (int i = hidden_Layers.Count - 1; i >= 0; i--) {
				if (hidden_Layers[i].inputConnection == null) {
					break;
				}

				Matrix prevConnTransposed = Matrix.Transpose(hidden_Layers[i].outputConnection.getConnectionMatrix);
				Matrix previousError = prevConnTransposed * currentLayerError;

				Matrix currGradient = Matrix.Map(prevConnTransposed, DeriveSigmoid);
				currGradient *= previousError;
				currGradient *= learningRate;

				Matrix prevStepMatrix = Matrix.Transpose(hidden_Layers[i].inputConnection.from.storage);
				Matrix prevErrorDelta = currGradient * prevStepMatrix;

				hidden_Layers[i].inputConnection.AdjustByDelta(prevErrorDelta);
				hidden_Layers[i].inputConnection.AdjustBias(currGradient);

				currentLayerError = previousError;
				//loop until the distribution happens between input and first hidden(the error from first hidden) adjust those
			}
			// now we have to adjust connections from input layer to the first hidden layer
			// NO the loop goes to the first hidden and looks at its input connections which is the output connection of our input layer
			//and adjusts it, inputs of input layer cannot be adjusted!		 
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
