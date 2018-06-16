using System.Collections.Generic;
using System.Speech.Recognition;

namespace Metin2SpeechToData {
	public class DefinitionParserData {
		/// <summary>
		/// Name of this definition --> file name that was used to parse this data
		/// </summary>
		public string ID { get; }

		/// <summary>
		/// Does this grammar have a companion Mob_'ID'.definition file
		/// </summary>
		public bool hasEnemyCompanionGrammar { get; set; }

		/// <summary>
		/// Groups defined at the top of the file
		/// </summary>
		public string[] groups { get; private set; }

		/// <summary>
		/// All items that are described in the file
		/// </summary>
		public Item[] entries { get; private set; }

		/// <summary>
		/// Grammar created from all the item names and ambiguities
		/// </summary>
		public Grammar grammar { get; private set; }

		public DefinitionParserData(string ID, string[] groups, Item[] entries) {
			this.ID = ID;
			this.groups = groups;
			this.entries = entries;
		}

		public struct Item {
			public Item(string mainPronounciation, string[] ambiguous, uint yangValue, string group) {
				this.mainPronounciation = mainPronounciation;
				this.ambiguous = ambiguous;
				this.yangValue = yangValue;
				this.group = group;
			}

			/// <summary>
			/// This items main pronounciation
			/// </summary>
			public string mainPronounciation { get; }

			/// <summary>
			/// All the ways you can call this item
			/// </summary>
			public string[] ambiguous { get; }

			/// <summary>
			/// How much is this item worth
			/// </summary>
			public uint yangValue { get; }

			/// <summary>
			/// Which group this item belongs to
			/// </summary>
			public string group { get; }
		}

		/// <summary>
		/// Gets main item pronounciation by comparing ambiguities
		/// </summary>
		public string GetMainPronounciation(string calledAmbiguity) {
			foreach (Item entry in entries) {
				if (entry.mainPronounciation == calledAmbiguity) {
					return calledAmbiguity;
				}
				foreach (string ambiguity in entry.ambiguous) {
					if (ambiguity == calledAmbiguity) {
						return entry.mainPronounciation;
					}
				}
			}
			if (!Configuration.debug) {
				throw new CustomException("This error should never be called becuse this function is called only when valid word from grammar is said!");
			}
			return null;
		}


		/// <summary>
		///	Returns parsed yang value for item 'itemName' itemName must be the main pronounciation!
		/// </summary>
		public uint GetYangValue(string itemName) {
			foreach (Item entry in entries) {
				if (entry.mainPronounciation == itemName) {
					return entry.yangValue;
				}
			}
			throw new CustomException("Data parsed incorrectly");
		}

		public void ConstructGrammar() {
			Choices main = new Choices();
			foreach (Item e in entries) {
				main.Add(e.mainPronounciation);
				foreach (string s in e.ambiguous) {
					main.Add(s);
				}
			}
			grammar = new Grammar(main) {
				Name = ID
			};
		}

		public string GetGroup(string itemName) {
			foreach (Item entry in entries) {
				if (entry.mainPronounciation == itemName) {
					return entry.group;
				}
			}
			throw new CustomException("Item doesn't exist in the entries, perhaps the archives are incomplete");
		}

		public Item GetItemEntry(string itemName) {
			foreach (Item item in entries) {
				if(item.mainPronounciation == itemName) {
					return item;
				}
			}
			throw new CustomException("Main pronounciation for " + itemName + " not found in entries");
		}

		public void AddItemDuringRuntime(Item item) {
			List<Item> itemList = new List<Item>(entries);
			foreach (Item entry in entries) {
				if(entry.mainPronounciation == item.mainPronounciation) {
					throw new CustomException("You can't add an item that already exists");
				}
			}
			itemList.Add(item);
			entries = itemList.ToArray();

			List<string> groupsList = new List<string>(groups);
			if (!groupsList.Contains(item.group)) {
				groupsList.Add(item.group);
				groups = groupsList.ToArray();
			}
			ConstructGrammar();
		}
	}
}
