using HouseInfo.ExcelHandler.Model.Abstract;

namespace HouseInfo.ExcelHandler.Model
{
    public class MatchingRange : IMatchingRange
    {
        public IRange FromRange { get; set; }
        public IRange ToRange { get; set; }

        public MatchingRange()
        {
        }
        
        public MatchingRange(IRange fromRange, IRange toRange)
        {
            FromRange = fromRange;
            ToRange = toRange;
        }
    }
}
