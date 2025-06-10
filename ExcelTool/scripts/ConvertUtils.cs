
public class ConvertUtils
{
  
  public static string GetLoadFieldMethod(string fieldType, int index)
    {
        if (string.IsNullOrEmpty(fieldType))
        {
            return $"ConvertUtils.GetString(fields[{index}])"; // 默认类型为 string
        }

        switch (fieldType)
        {
            case "int":
                return $"ConvertUtils.GetInt32(fields[{index}])";
            case "float":
                return $"ConvertUtils.GetFloat(fields[{index}])";
            case "double":
                return $"ConvertUtils.GetDouble(fields[{index}])";
            case "string":
                return $"ConvertUtils.GetString(fields[{index}])";
            case "bool":
                return $"ConvertUtils.GetBool(fields[{index}])";
            case "long":
                return $"ConvertUtils.GetLong(fields[{index}])";
            case "DateTime":
                return $"ConvertUtils.GetDateTime(fields[{index}])";
            default:
                return $"ConvertUtils.GetString(fields[{index}])"; // 默认类型为 string
        }
    }





    public static string GetType(string fieldType, string fieldName)
    {
        if (string.IsNullOrEmpty(fieldType))
        {
            return "string"; // 默认类型为 string
        }

        return fieldType.ToLowerInvariant() switch
        {
            "int" => "int",
            "float" => "float",
            "double" => "double",
            "string" => "string",
            "bool" => "bool",
            "long" => "long",
            "datetime" => "DateTime",
            _ => "string" // 默认类型为 string
        };
    }



    //=========================================================

    public static Int32 GetInt32(string data)
    {
        if (string.IsNullOrEmpty(data))
        {
            return 0;
        }

        try
        {
            return Convert.ToInt32(data);
        }
        catch (Exception e)
        {
            Console.WriteLine($"[读表]转换 Error int:{data}\n{e}");
        }
        return 0;
    }

    public static string GetString(string data)
    {
        if (string.IsNullOrEmpty(data))
        {
            return string.Empty;
        }
        data = data.Trim('"');
        return Convert.ToString(data);
    }

    public static float GetFloat(string data)
    {
        if (string.IsNullOrEmpty(data))
        {
            return 0f;
        }
         try
        {
           return Convert.ToSingle(data);
        }
        catch (Exception e)
        {
            Console.WriteLine($"[读表]转换 Error float:{data}\n{e}");
        }
       return 0f;
        
    }

    public static Double GetDouble(string data)
    {
        if (string.IsNullOrEmpty(data))
        {
            return 0d;
        }
        try
        {
            return Convert.ToDouble(data);
        }
        catch (Exception e)
        {
            Console.WriteLine($"[读表]转换 Error double:{data}\n{e}");
        }
       return 0d;
    }

    public static bool GetBool(string data)
    {
        if (string.IsNullOrEmpty(data))
        {
            return false;
        }

        try
        {
           return Convert.ToBoolean(Convert.ToInt32(data));
        }
        catch (Exception e)
        {
            Console.WriteLine($"[读表]转换 Error bool:{data}\n{e}");
        }
        return false;
    }

    public static DateTime GetDateTime(string data)
    {
        if (string.IsNullOrEmpty(data))
        {
            return DateTime.MinValue;
        }

        try
        {
            return Convert.ToDateTime(data);
        }
        catch (Exception e)
        {
            Console.WriteLine($"[读表]转换 Error DateTime:{data}\n{e}");
        }
        return DateTime.MinValue;
    }



    //====================================================================

    private static List<T> ConvertToListFromStr<T>(string data) where T : struct
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

        string[] strArray = data.Split(',');
        if (null == strArray || 0 == strArray.Length)
        {
            return new List<T>();
        }

        List<T> returnArray = new List<T>();
        foreach (var item in strArray)
        {
            returnArray.Add((T)Convert.ChangeType(item, typeof(T)));
        }

        return returnArray;
    }
        public static List<Int32> GetIntList(string data)
        {
            return ConvertToListFromStr<Int32>(data);
        }

        public static List<bool> GetBoolList(string data)
        {
            return ConvertToListFromStr<bool>(data);
        }

        public static List<float> GetFloatList(string data)
        {
            return ConvertToListFromStr<float>(data);
        }

        public static List<double> GetDoubleList(string data)
        {
            return ConvertToListFromStr<double>(data);
        }

        public static List<string> GetStringList(string data)
        {
            if (string.IsNullOrEmpty(data))
            {
                return new List<string>();
            }

            data = data.Trim('"');
            string[] strArray = data?.Split(',');
            if (null == strArray || 0 == strArray.Length)
            {
                return new List<string>();
            }

            return strArray.ToList();
        }



}