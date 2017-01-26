namespace HouseInfo.ExcelHandler
{
    public interface ITimeTest
    {
        double TotalTimeMs { get; }
        int AverageTime { get; }
        void Start(int iteration = 0);
        void Stop(int iteration = 0);
        string Save(string fileName);
    }

}