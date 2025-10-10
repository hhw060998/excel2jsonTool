using System;
using System.IO;
using Newtonsoft.Json;
using UnityEngine;

namespace GameUtility.IO
{
    public static class JsonSerializeUtility
    {
        public static T DeserializeFromFile<T>(string fullPath)
        {
            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException($"无法找到文件：{fullPath}");
            }

            try
            {
                var jsonText = File.ReadAllText(fullPath);
                var result = JsonConvert.DeserializeObject<T>(jsonText);
                return result;
            }
            catch (JsonSerializationException e)
            {
                throw new JsonSerializationException($"反序列化文件失败：{fullPath}", e);
            }
            catch (JsonException e)
            {
                throw new JsonException($"JSON格式错误：{fullPath}", e);
            }
            catch (Exception e)
            {
                throw new Exception($"发生未知错误：{fullPath}", e);
            }
        }

        public static void SerializeToFile<T>(string fullPath, T obj)
        {
            if (string.IsNullOrEmpty(fullPath))
            {
                Debug.LogError("尝试序列化的json文件路径为空");
                return;
            }

            if (obj == null)
            {
                Debug.LogError("尝试序列化的对象为空");
                return;
            }

            try
            {
                var setting = new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                };
                
                var jsonText = JsonConvert.SerializeObject(obj, setting);
                File.WriteAllText(fullPath, jsonText);
                Debug.Log($"序列化数据保存成功：{fullPath}");
            }
            catch (JsonSerializationException e)
            {
                throw new JsonSerializationException($"序列化对象失败：{fullPath}", e);
            }
            catch (JsonException e)
            {
                throw new JsonException($"JSON格式错误：{fullPath}", e);
            }
            catch (Exception e)
            {
                throw new Exception($"发生未知错误：{fullPath}", e);
            }
        }
    }
}