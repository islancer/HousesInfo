using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HouseInfo.ExcelHandler
{
    [Serializable]
    public class TimeTest : ITimeTest
    {
        private DateTime _timeBegin;
        private DateTime _timeEnd;
        private int _currentIteration;
        public double TotalTimeMs
        {
            get
            {
                return AllTimeTests.Keys.Sum(key => AllTimeTests[key]);
            }
        }

        

        public int AverageTime {
            get { return (int) (TotalTimeMs/AllTimeTests.Count); }
        }

        [NonSerialized] public IDictionary<int, double> AllTimeTests;

        

        public TimeTest()
        {
            AllTimeTests = new Dictionary<int, double>();
            _currentIteration = -1;
            
        }

        public void Start(int currentIteration = -1)
        {
            if (currentIteration >= 0)
            {
                _timeBegin = DateTime.Now;
                _currentIteration = currentIteration;
            }
        }

        public void Stop(int iteration = -1)
        {
            if (_currentIteration >= 0)
            {
                _timeEnd = DateTime.Now;
                double timePerIteration = (new TimeSpan(_timeEnd.Ticks - _timeBegin.Ticks)).TotalMilliseconds;
                AllTimeTests.Add(_currentIteration, timePerIteration);
                _currentIteration = -1;
            }
        }

        public string Save(string fileName)
        {
            try
            {
                using (StreamWriter writer = File.CreateText(fileName))
                {
                    writer.WriteLine("Общее время:   {0:HH:mm:ss}", (new DateTime()).AddMilliseconds(TotalTimeMs));
                    writer.WriteLine("Общее время:   {0} мс", TotalTimeMs);
                    writer.WriteLine("Время за итерацию: {0}", AverageTime);
                    writer.WriteLine("Всего итераций:    {0}",  AllTimeTests.Count);

                    foreach (var key in AllTimeTests.Keys)
                    {
                        writer.WriteLine("итерация: {0:0000}    время  {1:####} мс", key, AllTimeTests[key]);
                    }
                    writer.Flush();
                    writer.Close();
                }
                return null;

            }
            catch (Exception ex)
            {                
                return ex.Message;            
            }
        }
    }
}
