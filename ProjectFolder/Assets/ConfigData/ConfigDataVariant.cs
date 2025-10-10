using System;

namespace Data
{
    public abstract class ConfigDataWithKey<TConfigInfo, TEnum> : ConfigDataBase<TConfigInfo>, IConfigDataWithKey<TConfigInfo>
        where TEnum : Enum where TConfigInfo : IConfigRawInfo
    {
        public override TConfigInfo GetDataByKey<TEnum1>(TEnum1 key)
        {
            if (key is TEnum enumKey)
            {
                return GetData(Convert.ToInt32(enumKey));
            }

            throw new InvalidOperationException($"传入的key类型错误，期望类型:{typeof(TEnum).Name}，实际类型:{typeof(TEnum1).Name}");
        }
    }

    public interface IConfigDataWithKey<out T>
    {
        T GetDataByKey<TEnum>(TEnum key) where TEnum : Enum;
    }

    public abstract class ConfigDataWithCompositeId<TConfigInfo> : ConfigDataBase<TConfigInfo>
        where TConfigInfo : IConfigRawInfo
    {
        protected abstract int CompositeMultiplier { get; }

        public TConfigInfo GetDataByCompositeKey(int key1, int key2)
        {
            if (key1 < 0 || key1 >= CompositeMultiplier)
            {
                throw new ArgumentOutOfRangeException($"{nameof(key1)}的取值范围为0~{CompositeMultiplier - 1}，实际值:{key1}");
            }

            if (key2 < 0 || key2 >= CompositeMultiplier)
            {
                throw new ArgumentOutOfRangeException($"{nameof(key2)}的取值范围为0~{CompositeMultiplier - 1}，实际值:{key2}");
            }

            var compositeId = key1 * CompositeMultiplier + key2;
            return GetData(compositeId);
        }
    }
}