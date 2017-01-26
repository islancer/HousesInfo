using System.Collections.Generic;
using HouseInfo.ExcelHandler.Model.Abstract;

// ReSharper disable once CheckNamespace
namespace HouseInfo.ExcelHandler.Settings
{
    public interface ICalcSettings : IWorksheetSettings
    {
        IRange AddressRange { get; set; }
        IList<IMatchingRange> СopiedRangesToFlats { get; set; }
        IList<IMatchingRange> СopiedRangesToHouses { get; set; }
    }

    
}
