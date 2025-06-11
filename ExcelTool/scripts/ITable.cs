
namespace GameFramework.Table
{
    //=========================================================
    // Interface for table data loading
    //=========================================================

    public interface ITable
    {
        public void Load(string[] data);
        public int GetId();
    }
}
