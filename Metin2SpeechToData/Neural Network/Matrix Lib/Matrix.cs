using System;
using System.Collections.Generic;

namespace Metin2SpeechToData.Neural_Network {
	class Matrix {

		private bool isSquare { get; }
		private int cols { get; }
		private int rows { get; }

		private double[,] _matrix;

		public Matrix(int rows, int cols) {
			_matrix = new double[rows, cols];
			isSquare = rows == cols;
			this.cols = cols;
			this.rows = rows;
		}

		/// <summary>
		/// Randomly fills this matrix with values from 0 inclusive to 1.0 exlusive
		/// </summary>
		public void InitRandomOneNormalized() {
			Random r = new Random();
			for (int i = 0; i < rows; i++) {
				for (int j = 0; j < cols; j++) {
					_matrix[i, j] = r.NextDouble() * 2 - 1;
					Console.WriteLine(_matrix[i, j]);
				}
			}
			Console.ReadLine();
		}

		/// <summary>
		/// Makes a unit matrix out of this
		/// </summary>
		public void UnitMatrix() {
			for (int i = 0; i < rows; i++) {
				for (int j = 0; j < cols; j++) {
					if (i == j) {
						_matrix[i, j] = 1;
					}
					else {
						_matrix[i, j] = 0;
					}
				}
			}
		}

		public static Matrix operator +(Matrix a, Matrix b) {
			if (a.rows != b.rows || a.cols != b.cols) {
				throw new Exception("Attempting to sum two matrices with mismatched sizes.");
			}
			Matrix newM = new Matrix(a.rows, a.cols);
			for (int i = 0; i < a.rows; i++) {
				for (int j = 0; j < a.cols; j++) {
					newM._matrix[i,j] = a._matrix[i, j] + b._matrix[i, j];
				}
			}
			return newM;
		}

		public static Matrix operator -(Matrix a, Matrix b) {
			if (a.rows != b.rows || a.cols != b.cols) {
				throw new Exception("Attempting to sum two matrices with mismatched sizes.");
			}
			Matrix newM = new Matrix(a.rows, a.cols);
			for (int i = 0; i < a.rows; i++) {
				for (int j = 0; j < a.cols; j++) {
					newM._matrix[i,j] = a._matrix[i, j] - b._matrix[i, j];
				}
			}
			return newM;
		}

		public static Matrix FromAray(double[] data) {
			Matrix m = new Matrix(data.Length, 1);
			for (int i = 0; i < data.Length; i++) {
				m.raw[i, 0] = data[i];
			}
			return m;
		}

		public double[] ToArray() {
			double[] array = new double[rows];
			for (int i = 0; i < array.Length; i++) {
				array[i] = _matrix[i, 0];
			}
			return array;
		}

		public static Matrix operator *(Matrix a, Matrix b) {
			//If A is an m - by - n matrix and B is an n - by - p matrix,
			// then their matrix product AB is the m - by - p matrix whose entries are given by dot product
			// of the corresponding row of A and the corresponding column of B:
			Matrix x = new Matrix(a.rows, b.cols);

			for (int i = 0; i < x.rows; i++) {
				for (int j = 0; j < x.cols; j++) {
					double value = 0;
					for (int k = 0; k < a.cols; k++) {
						value += (a._matrix[i, k] * b._matrix[k, j]);
					}
					x._matrix[i, j] = value;
				}
			}
			return x;
		}

		public static Matrix operator *(Matrix a, double b) {
			Matrix x = new Matrix(a.rows, a.cols);

			for (int i = 0; i < x.rows; i++) {
				for (int j = 0; j < x.cols; j++) {
					x._matrix[i, j] = a._matrix[i,j] * b;
				}
			}
			return x;
		}

		/// <summary>
		/// Get new transposed matrix from the old one, new instance is returned.
		/// </summary>
		public static Matrix Transpose(Matrix m) {
			Matrix x = new Matrix(m.cols, m.rows);

			for (int i = 0; i < m.rows; i++) {
				for (int j = 0; j < m.cols; j++) {
					x._matrix[j, i] = m._matrix[i, j];
				}
			}
			return x;
		}

		/// <summary>
		/// Pront matrix data in a readable format
		/// </summary>
		public static void Print(Matrix m) {
			for (int i = 0; i < m.rows; i++) {
				for (int j = 0; j < m.cols; j++) {
					Console.Write(m._matrix[i, j] + " ");
				}
				Console.WriteLine();
			}
			Console.WriteLine();
		}

		/// <summary>
		/// Modify each value of the matrix with this function
		/// </summary>
		public void Map(Func<double, double> mapperFunction) {
			for (int i = 0; i < rows; i++) {
				for (int j = 0; j < cols; j++) {
					_matrix[i, j] = mapperFunction(_matrix[i, j]);
				}
			}
		}

		/// <summary>
		/// Modify each value of the matrix with this function and return new
		/// </summary>
		public static Matrix Map(Matrix m, Func<double, double> mapperFunction) {
			Matrix newM = new Matrix(m.rows, m.cols);
			for (int i = 0; i < newM.rows; i++) {
				for (int j = 0; j < newM.cols; j++) {
					newM._matrix[i, j] = mapperFunction(m._matrix[i, j]);
				}
			}
			return newM;
		}

		/// <summary>
		/// Raw matrix 2D array
		/// </summary>
		public double[,] raw {
			get { return _matrix; }
		}
	}
}
