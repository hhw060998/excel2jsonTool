using System;
using System.Collections.Generic;
using System.Linq;

namespace Data.TableScript
{
	/// <summary> This is auto-generated, don't modify manually </summary>
	public class SampleCombinedKeyInfo
	{
		/// <summary>
		/// 编号: 组合key仅支持int类型，且不能超过46340
		/// </summary>
		public int id { get; set; }
		
		/// <summary>
		/// 字符串类型
		/// </summary>
		public int group { get; set; }
		
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
	public class SampleCombinedKeyConfig
	{
		private static Dictionary<int, SampleCombinedKeyInfo> _data;
		
		private const int COMPOSITE_MULTIPLIER = 46340;
		
		
		public static void Initialize()
		{
			_data = ConfigDataUtility.DeserializeConfigData<Dictionary<int, SampleCombinedKeyInfo>>(nameof(SampleCombinedKeyConfig));
		}
		
		public static SampleCombinedKeyInfo GetDataById(int id)
		{
			if(_data.TryGetValue(id, out var result))
			{
				return result;
			}
			throw new InvalidOperationException($"Can not find the config data by id: {id}.");
		}
		
		public static int CombineKey(int key1, int key2)
		{
			if (key1 is < 0 or >= COMPOSITE_MULTIPLIER) throw new ArgumentOutOfRangeException(nameof(key1));
			if (key2 is < 0 or >= COMPOSITE_MULTIPLIER) throw new ArgumentOutOfRangeException(nameof(key2));
			return key1 * COMPOSITE_MULTIPLIER + key2;
		}
		
		public static SampleCombinedKeyInfo GetDataByCompositeKey(int id, int group)
		{
			// Use combined key generated from (id,group)
			return GetDataById(CombineKey(id, group));
		}
		
		public static IEnumerable<TResult> SelectValueCollection<TResult>(Func<SampleCombinedKeyInfo, TResult> selector)
		{
			return _data.Values.Select(selector);
		}
		
		public static IEnumerable<SampleCombinedKeyInfo> GetInfoCollection(Func<SampleCombinedKeyInfo, bool> predicate)
		{
			return _data.Values.Where(predicate);
		}
	}
}