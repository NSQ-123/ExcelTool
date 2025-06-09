
public class ConvertUtils
{
    public static object ConvertField(string field, Type targetType)
    {
        if (string.IsNullOrEmpty(field))
        {
            // 如果字段为空，返回目标类型的默认值
            return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;
        }
        
        try
        {
            if (targetType == typeof(int))
            {
                return int.TryParse(field, out var result) ? result : 0;
            }
            if (targetType == typeof(float))
            {
                return float.TryParse(field, out var result) ? result : 0f;
            }
            if (targetType == typeof(double))
            {
                return double.TryParse(field, out var result) ? result : 0.0;
            }
            if (targetType == typeof(bool))
            {
                return bool.TryParse(field, out var result) ? result : false;
            }
            if (targetType == typeof(long))
            {
                return long.TryParse(field, out var result) ? result : 0L;
            }
            if (targetType == typeof(DateTime))
            {
                return DateTime.TryParse(field, out var result) ? result : DateTime.MinValue;
            }
            if (targetType == typeof(string))
            {
                return field;
            }
        }
        catch
        {
            // 如果转换失败，返回目标类型的默认值
            return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;
        }

        // 默认返回字符串
        return field;
    }

    public static T ConvertField<T>(string field)
    {
        if (string.IsNullOrEmpty(field))
        {
            // 如果字段为空，返回目标类型的默认值
            return default(T);
        }
        
        try
        {
           
        }
        catch
        {
            
        }

        // 默认返回字符串
         return default(T);
    }





    public static Type GetType(string fieldType)
    {
        if (string.IsNullOrEmpty(fieldType))
        {
            return typeof(string); // 默认类型为 string
        }

        return fieldType.ToLowerInvariant() switch
        {
            "int" => typeof(int),
            "float" => typeof(float),
            "double" => typeof(double),
            "string" => typeof(string),
            "bool" => typeof(bool),
            "long" => typeof(long),
            "datetime" => typeof(DateTime),
            _ => typeof(string) // 默认类型为 string
        };
    }



    // public static string NormalizeFieldType(string fieldType)
    // {
    //     if (string.IsNullOrEmpty(fieldType))
    //     {
    //         return fieldType;
    //     }

    //     fieldType = fieldType.ToLowerInvariant() switch
    //     {
    //         "int" => "int",
    //         "float" => "float",
    //         "double" => "double",
    //         "string" => "string",
    //         "bool" => "bool",
    //         "long" => "long",
    //         "datetime" => "DateTime",
    //         _ => "string" // 默认类型为 string
    //     };
    //     return fieldType;
    // }



}