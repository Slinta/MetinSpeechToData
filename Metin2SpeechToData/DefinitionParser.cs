using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Metin2SpeechToData {
	public class DefinitionParser {

		private Dictionary<string, string[]> availableDefinitions;

		public DefinitionParser() {
			DirectoryInfo d = new DirectoryInfo(Directory.GetCurrentDirectory());
			FileInfo[] filesPresent = d.GetFiles("*.definition");

			if (filesPresent.Length == 0) {
				throw new Exception("Your program is missing voice recognition strings! Either redownload, or create your own *.definition text file.");
			}

			availableDefinitions = new Dictionary<string, string[]>();
			for (int i = 0; i < availableDefinitions.Count; i++) {
				List<string> strings = new List<string>();
				using (StreamReader s = filesPresent[i].OpenText()) {
					while (!s.EndOfStream) {
						strings.Add(s.ReadLine());
					}
				}
				availableDefinitions.Add(filesPresent[i].Name, strings.ToArray());
			}
		}

		public string[] getSprcificDefinitions(string identifier) {
			return availableDefinitions[identifier];
		}
	}
}
