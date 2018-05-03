using System;
using System.Collections.Generic;

namespace Metin2SpeechToData.Neural_Network {
	class Matrix {

		private bool isSquare { get; }
		private int _width { get; }
		private int _height { get; }

		private double[,] _matrix;

		public Matrix(int width, int height) {
			_matrix = new double[width, height];
			isSquare = width == height;
			_width = width;
			_height = height;
		}

		/// <summary>
		/// Randomly fills this matrix with values from 0 inclusive to 1.0 exlusive
		/// </summary>
		public void InitRandomOneNormalized() {
			Random r = new Random();
			for (int i = 0; i < _width; i++) {
				for (int j = 0; j < _height; j++) {
					_matrix[i, j] = r.NextDouble();
				}
			}
		}

		/// <summary>
		/// Makes a unit matrix out of this
		/// </summary>
		public void UnitMatrix() {
			for (int i = 0; i < _width; i++) {
				for (int j = 0; j < _height; j++) {
					if(i == j) {
						_matrix[i, j] = 1;
					}
					else {
						_matrix[i, j] = 0;
					}
				}
			}
		}


		public static Matrix operator +(Matrix a, Matrix b) {
			if (a._height != b._height || a._width != b._width) {
				throw new Exception("Attempting to sum two matrices with mismatched sizes.");
			}
			for (int i = 0; i < a._width; i++) {
				for (int j = 0; j < a._height; j++) {
					a._matrix[i, j] += b._matrix[i, j];
				}
			}
			return a;
		}


		public static Matrix operator *(Matrix a, Matrix b) {
			//If A is an m - by - n matrix and B is an n - by - p matrix,
			// then their matrix product AB is the m - by - p matrix whose entries are given by dot product
			// of the corresponding row of A and the corresponding column of B:
			Matrix x = new Matrix(a._width, b._height);

			for (int i = 0; i < a._width; i++) {
				for (int j = 0; j < b._height; j++) {
					x._matrix[i, j] = MatrixMultiplication(a, b, i, j);
				}
			}
			return x;
		}

		private static double MatrixMultiplication(Matrix a, Matrix b, int row, int col) {
			double value = 0;
			for (int j = 0; j < b._height; j++) {
				value += (a._matrix[row, j] * b._matrix[j, col]);
			}
			return value;
		}

		/// <summary>
		/// Get new transposed matrix from the old one, new instance is returned.
		/// </summary>
		public static Matrix Transpose(Matrix m) {
			Matrix x = new Matrix(m._height, m._width);

			for (int i = 0; i < m._width; i++) {
				for (int j = 0; j < m._height; j++) {
					x._matrix[j, i] = m._matrix[i, j];
				}
			}
			return x;
		}

		/// <summary>
		/// Pront matrix data in a readable format
		/// </summary>
		public static void Print(Matrix m) {
			for (int i = 0; i < m._width; i++) {
				for (int j = 0; j < m._height; j++) {
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
			for (int i = 0; i < _width; i++) {
				for (int j = 0; j < _height; j++) {
					_matrix[i, j] = mapperFunction(_matrix[i, j]);
				}
			}
		}

		/// <summary>
		/// Raw matrix 2D array
		/// </summary>
		public double[,] raw {
			get { return _matrix; }
		}
	}
}
