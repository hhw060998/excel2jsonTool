using System.Reflection;
using UnityGameFramework.Runtime;
using System;
using System.IO;
using Newtonsoft.Json;
using UnityEngine;

namespace Data
{
    public static class ConfigDataUtility
    {
        private static readonly string JsonPath = $"{Application.streamingAssetsPath}/Json/";
        private const string AssemblyName = "GameMain.Runtime";

        public static T DeserializeConfigData<T>(string fileName) where T : class
        {
            var filePath = Path.Combine(JsonPath, $"{fileName}.json");
            try
            {
                var json = File.ReadAllText(filePath);
                var data = JsonConvert.DeserializeObject<T>(json);
                if (data != null)
                {
                    Log.Debug(filePath + "加载成功");
                    return data;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"{filePath}加载失败：{ex.Message}");
            }

            throw new Exception($"无法加载文件：{filePath}, 必须确保Json名为{nameof(T)}");
        }

        /// <summary>
        /// 初始化所有的表格数据类，需要确保类中包含InitializeAll静态方法
        /// </summary>
        public static void InitializeAll()
        {
            var assembly = Assembly.Load(AssemblyName);
            if (assembly == null)
            {
                throw new NullReferenceException($"Assembly [{AssemblyName}] load failed, can't load config data");
            }
            var types = assembly.GetTypes();

            foreach (var type in types)
            {
                if (!type.IsClass || type.IsAbstract) continue;
                if (type.GetAttribute<ConfigDataAttribute>() != null)
                {
                    var initializeMethod = type.GetMethod("Initialize", BindingFlags.Static | BindingFlags.Public);
                    if (initializeMethod != null)
                    {
                        initializeMethod.Invoke(null, null);
                    }
                    else
                    {
                        Log.Error($"找不到{type.Name}的Initialize方法");
                    }
                }
            }
        }
    }
}