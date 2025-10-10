using System;
using System.Collections.Generic;

namespace Data
{
    public interface IConfigDataBase
    {
        /// <summary>按 id 获取单条，返回非泛型接口类型。</summary>
        IConfigRawInfo GetData(int id);

        /// <summary>按谓词获取集合，谓词以 IConfigRawInfo 为参数。</summary>
        IEnumerable<IConfigRawInfo> GetCollection(Func<IConfigRawInfo, bool> predicate);

        /// <summary>将值集合投影为 TResult（selector 接受 IConfigRawInfo）。</summary>
        IEnumerable<TResult> SelectValueCollection<TResult>(Func<IConfigRawInfo, TResult> selector);

        /// <summary>加载数据（由外部统一调用）。</summary>
        void Load();

        /// <summary>获取用于显示的枚举映射（如果有）。</summary>
        Dictionary<int, string> GetDisplayMap();
    }
}