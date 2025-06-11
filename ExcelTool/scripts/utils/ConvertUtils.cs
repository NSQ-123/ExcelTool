
namespace GameFramework.Table
{
    public class ConvertUtils
    {
        public static T Get<T>(string data) where T : IConvertible
        {
            if (string.IsNullOrEmpty(data)) return default(T);

            try
            {
                return (T)Convert.ChangeType(data, typeof(T));
            }
            catch (Exception e)
            {
                Console.WriteLine($"[读表]转换 Error {typeof(T).Name}:{data}\n{e}");
            }
            return default(T);
        }
        
        public static List<T> GetList<T>(string data)where T : IConvertible
        {
            if (string.IsNullOrEmpty(data))
            {
                return new List<T>();
            }

            data = data.Trim('"');
            if (data[0] == '\"')
            {
                data = data.Substring(1, data.Length - 2);
            }

            var strArray = data.Split(',');
            return 0 == strArray.Length ? new List<T>() : GetList<T>(strArray);
        }

        public static List<T> GetList<T>(string[] data)where T : IConvertible
        {
            if(data ==null || data.Length == 0)
            {
                return new List<T>();
            }
            
            List<T> result = new List<T>();
            foreach (var item in data)
            {
                result.Add((T)Convert.ChangeType(item, typeof(T)));
            }
            return result;
        }
        
        public static List<T> LoadArr<T>(string content) where T : ITable, new()
        {
            content = content.Trim('"');
            List<T> list = new List<T>();
            var rows = content.Split(';');
            for (var i = 0; i < rows.Length; i++)
            {
                if (rows[i].Length > 0)
                {
                    string[] rowValues = rows[i].Split(',');
                    T t = new T();
                    t.Load(rowValues);
                    list.Add(t);
                }
            }
            return list;
        }

    }
}