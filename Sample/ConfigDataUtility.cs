using System.Reflection;
using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using UnityEngine;

namespace Data
{
    public static class ConfigDataUtility
    {
        private static readonly string JsonPath = Path.Combine(Application.streamingAssetsPath, "Json");
        private const string AssemblyName = "GameMain.Runtime";

        public static T DeserializeConfigData<T>(string fileName) where T : class
        {
            var filePath = Path.Combine(JsonPath, $"{fileName}.json");
            
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"无法找到文件：{filePath}");
            }

            try
            {
                var json = File.ReadAllText(filePath);
                var data = JsonConvert.DeserializeObject<T>(json);
                
                if (data == null)
                {
                    throw new InvalidOperationException($"文件内容无效：{fileName}");
                }

                Debug.Log($"{filePath} 加载成功");
                return data;
            }
            catch (JsonException jsonEx)
            {
                throw new InvalidOperationException($"文件格式错误：{fileName}, 错误详情: {jsonEx.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"{filePath} 加载失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 初始化所有的表格数据类，需要确保类中包含 Initialize 静态方法
        /// </summary>
        public static void InitializeAll()
        {
            Assembly assembly;
            try
            {
                assembly = Assembly.Load(AssemblyName);
            }
            catch (FileNotFoundException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception($"加载程序集失败: {ex.Message}");
            }

            var types = assembly.GetTypes();

            foreach (var type in types.Where(t => t.IsClass && !t.IsAbstract))
            {
                if (Attribute.IsDefined(type, typeof(ConfigDataAttribute)))
                {
                    var initializeMethod = type.GetMethod("Initialize", BindingFlags.Static | BindingFlags.Public);
                    
                    if (initializeMethod != null)
                    {
                        try
                        {
                            initializeMethod.Invoke(null, null);
                            Debug.Log($"{type.Name} 初始化成功");
                        }
                        catch (TargetInvocationException ex)
                        {
                            Debug.LogError($"调用 {type.Name}.Initialize 失败: {ex.InnerException?.Message}");
                        }
                    }
                    else
                    {
                        Debug.LogWarning($"{type.Name} 没有找到 Initialize 静态方法");
                    }
                }
            }
        }
    }
}
