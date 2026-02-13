using System.IO;
using System.Runtime.Serialization.Json;

namespace AutoOpener.Core.IO
{
    public static class JsonStorage
    {
        public static void Write<T>(string filePath, T obj)
        {
            var ser = new DataContractJsonSerializer(typeof(T));
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                ser.WriteObject(fs, obj);
        }

        public static T Read<T>(string filePath)
        {
            var ser = new DataContractJsonSerializer(typeof(T));
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                return (T)ser.ReadObject(fs);
        }
    }
}
