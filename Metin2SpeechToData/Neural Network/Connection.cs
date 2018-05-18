using System;
namespace Metin2SpeechToData.Neural_Network {
	public sealed class Connection {

		public Layer from { get; }
		public Layer to { get; }

		public Connection(Layer from, Layer to, int seed) {
			this.from = from;
			this.to = to;
			getConnectionMatrix = new Matrix(to.neuronCount, from.neuronCount);
			getConnectionMatrix.InitRandomOneNormalized(seed * Environment.TickCount);
			getBias = new Matrix(to.neuronCount, 1);
			getBias.InitRandomOneNormalized(seed);
		}

		public Matrix getConnectionMatrix { get; private set; }

		public Matrix getBias { get; private set; }

		public void AdjustByDelta(Matrix prev_errorDelta) {
			getConnectionMatrix += prev_errorDelta;
		}

		internal void AdjustBias(Matrix gradient) {
			getBias += gradient;
		}
	}
}
