using System;
using System.Collections.Generic;

public partial class T_Person
{
    private static Dictionary<int, T_Person> _dataDic = new Dictionary<int, T_Person>();
    private static List<T_Person> _dataList;

    /// <summary>
    /// 标识符
    /// </summary>
    public int ID { get; set; }
    /// <summary>
    /// 名字
    /// </summary>
    public string Name { get; set; }
    /// <summary>
    /// 年龄
    /// </summary>
    public int Age { get; set; }

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
}
