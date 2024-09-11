using System;
using System.Collections.Generic;
using System.Linq;

namespace Data.TableScript
{
	/// <summary> This is auto-generated, don't modify manually </summary>
	public class SampleInfo
	{
		public int id { get; set; }
		public string name { get; set; }
		public int damage { get; set; }
		public float damage_increase { get; set; }
		public List<string> attr_names { get; set; }
		public List<int> attr_values { get; set; }
		public List<float> float_values { get; set; }
		public Dictionary<int,string> dict_values1 { get; set; }
		public Dictionary<string,float> dict_values2 { get; set; }
	}
	
	[ConfigData]
	public class SampleConfig
	{
		private static Dictionary<int, SampleInfo> _data;
		
		public static void Initialize()
		{
			_data = ConfigDataUtility.DeserializeConfigData<Dictionary<int, SampleInfo>>(nameof(SampleConfig));
		}
		
		public static SampleInfo GetDataById(int id)
		{
			if(_data.TryGetValue(id, out var result))
			{
				return result;
			}
			throw new InvalidOperationException();
		}
		
		
		
		public static IEnumerable<TResult> SelectValueCollection<TResult>(Func<SampleInfo, TResult> selector)
		{
			return _data.Values.Select(selector);
		}
		
		public static IEnumerable<SampleInfo> GetInfoCollection(Func<SampleInfo, bool> predicate)
		{
			return _data.Values.Where(predicate);
		}
	}
}