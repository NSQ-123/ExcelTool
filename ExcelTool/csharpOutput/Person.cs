using System;
using System.Collections.Generic;

public partial class T_Person
{
    private static Dictionary<int, T_Person> _dataDic = new Dictionary<int, T_Person>();
    private static List<T_Person> _dataList;

    /// <summary>
    /// 标识符
    /// </summary>
    public System.Int32 ID { get; set; }
    /// <summary>
    /// 名字
    /// </summary>
    public System.String Name { get; set; }
    /// <summary>
    /// 年龄
    /// </summary>
    public System.Int32 Age { get; set; }

    public static T_Person GetById(int id)
    {
        if (_dataDic.TryGetValue(id, out var value))
        {
            return value;
        }
        return null;
    }

    public static List<T_Person> GetAll()
    {
        if (_dataList == null)
        {
            _dataList = new List<T_Person>(_dataDic.Values);
        }
        return _dataList;
    }

    public void Load(string csvline)
    {
       if (string.IsNullOrEmpty(csvline)) return;
       // 按逗号分隔字段
       var fields = csvline.Split(',');
       if (fields.Length < 1) return;
       // 给实例赋值
       this.ID =ConvertUtils.ConvertField<System.Int32>(fields[0]);

       this.Name =ConvertUtils.ConvertField<System.String>(fields[1]);

       this.Age =ConvertUtils.ConvertField<System.Int32>(fields[2]);


    }
}
