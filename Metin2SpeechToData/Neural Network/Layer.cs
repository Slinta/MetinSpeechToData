using System;
using System.Collections.Generic;

namespace Metin2SpeechToData.Neural_Network {
	public class Layer {

		public int layerIndex { get; }
		public int neuronCount { get; private set; }

		public Connection inputConnection;
		public Connection outputConnection;

		public Matrix storage;

		public Layer(int layerNeurons, int currentLayer) {
			layerIndex = currentLayer;
			neuronCount = layerNeurons;
		}

		public void ConnectLayer(Layer other, int seed) {
			Connection c = new Connection(this, other, seed);
			other.inputConnection = this.outputConnection = c;
		}

		public Matrix ProcessInput(Matrix matrix) {
			Matrix toHidden = outputConnection.getConnectionMatrix * matrix;
			toHidden += outputConnection.getBias;
			toHidden.Map(Sigmoid);
			//Console.WriteLine("From input " + layerIndex + " to hidden" + outputConnection.to.layerIndex + ": ");
			//Matrix.Print(toHidden);
			storage = toHidden;
			return outputConnection.to.ProcessLayers(toHidden);
		}

		private Matrix ProcessLayers(Matrix matrix) {
			if (outputConnection.to.outputConnection != null) {
				Matrix hiddenMatrix = outputConnection.getConnectionMatrix * matrix;
				hiddenMatrix += outputConnection.getBias;
				hiddenMatrix.Map(Sigmoid);
				//Console.WriteLine("From hidden " + layerIndex + " to hidden " + outputConnection.to.layerIndex);
				//Matrix.Print(hiddenMatrix);
				storage = hiddenMatrix;
				return outputConnection.to.ProcessLayers(hiddenMatrix);
			}
			else {
				Matrix outMatrix = outputConnection.getConnectionMatrix * matrix;
				outMatrix += outputConnection.getBias;
				outMatrix.Map(Sigmoid);
				//Console.WriteLine("From hidden " + layerIndex + " to output (layer " + outputConnection.to.layerIndex + ")");
				//Matrix.Print(outMatrix);
				storage = outMatrix;
				return outMatrix;
			}
		}

		private double Sigmoid(double x) {
			return 1 / (1 + Math.Exp(-x));
		}

		private double DeriveSigmoid(double x) {
			return x - (1 - x);
		}
	}
}
