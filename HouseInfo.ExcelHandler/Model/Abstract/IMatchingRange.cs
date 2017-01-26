
namespace HouseInfo.ExcelHandler.Model.Abstract
{
    public interface IMatchingRange
    {
        IRange FromRange { get; set; }
        IRange ToRange { get; set; }
    }
}
