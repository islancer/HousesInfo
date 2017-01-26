using HouseInfo.ExcelHandler.Model.Abstract;

namespace HouseInfo.ExcelHandler.Model
{
    public class Range : IRange
    {
        public string Begin { get; set; }

        public string End { get; set; }

        public Range()
        {
        }


        public Range(string beginEndRange)
        {
            Begin = beginEndRange;
            End = beginEndRange;
        }

        public Range(string beginRange, string endRange)
        {
            Begin = beginRange;
            End = endRange;
        }
    }  
}
