using System;
using System.Collections.Generic;
using System.Linq;

namespace Data.TableScript
{
	/// <summary> This is auto-generated, don't modify manually </summary>
	public class SampleInfo
	{
		/// <summary>
		/// 编号
		/// </summary>
		public int id { get; set; }
		
		/// <summary>
		/// 字符串类型
		/// </summary>
		public string name { get; set; }
		
		/// <summary>
		/// 整型
		/// </summary>
		public int damage { get; set; }
		
		/// <summary>
		/// 浮点型
		/// </summary>
		public float damage_increase { get; set; }
		
		/// <summary>
		/// 字符串列表
		/// </summary>
		public List<string> attr_names { get; set; }
		
		/// <summary>
		/// 整型列表
		/// </summary>
		public List<int> attr_values { get; set; }
		
		/// <summary>
		/// 浮点数列表
		/// </summary>
		public List<float> float_values { get; set; }
		
		/// <summary>
		/// 字典-整型:字符串
		/// </summary>
		public Dictionary<int,string> dict_values1 { get; set; }
		
		/// <summary>
		/// 字典-字符串:浮点
		/// </summary>
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
			throw new InvalidOperationException($"Can not find the config data by id: {id}.");
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