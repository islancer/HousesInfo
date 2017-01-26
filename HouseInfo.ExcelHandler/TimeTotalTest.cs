using System;
using System.IO;

namespace HouseInfo.ExcelHandler
{
    public class TimeTotalTest : ITimeTest
    {
        private DateTime _timeBegin;
        private DateTime _timeEnd;
        private int _beginIteration;
        private int _endIteration;

        public TimeTotalTest()
        {
            _beginIteration = -1;
            _endIteration = -1;           
            _timeBegin = new DateTime();
            _timeEnd = _timeBegin;
        }

        public double TotalTimeMs
        {
            get
            {              
                    return (new TimeSpan(_timeEnd.Ticks - _timeBegin.Ticks)).TotalMilliseconds;          
            }
        }

        public int AverageTime
        {
            get
            {
                return (Math.Max(_endIteration, _beginIteration) != _beginIteration)
                    ? (int) (TotalTimeMs/(_endIteration - _beginIteration))
                    : 0;
            }
        }



        public void Start(int iteration = 0)
        {
            if (iteration >= 0)
            {
                _timeBegin = DateTime.Now;
                _beginIteration = iteration;
            }
        }

        public void Stop(int iteration = 0)
        {
            if (_beginIteration > -1)
            {
                if (_beginIteration == iteration)
                    _endIteration++;
                else
                    _endIteration = iteration;

                if (_endIteration > _beginIteration)
                    _timeEnd = DateTime.Now;
            }

        }

        public string Save(string fileName)
        {
            try
            {
               // using (StreamWriter writer = File.CreateText(fileName))
                using (StreamWriter writer = File.AppendText(fileName))
                {                   
                    writer.WriteLine("_____________________________________________");
                    writer.WriteLine("Завершено в: {0} ", _timeEnd);
                    writer.WriteLine("-----");
                    writer.WriteLine("Общее время:   {0:HH:mm:ss}", (new DateTime()).AddMilliseconds(TotalTimeMs));
                    writer.WriteLine("Общее время:   {0} мс", TotalTimeMs);
                    writer.WriteLine("Время за итерацию: {0}", AverageTime);
                    writer.WriteLine("Всего итераций:    {0}", _endIteration - _beginIteration);
                    writer.WriteLine("");
                    
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
