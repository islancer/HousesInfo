 // ReSharper disable once CheckNamespace

using HouseInfo.ExcelHandler.Model.Abstract;

namespace HouseInfo.ExcelHandler.Settings
{
    public interface IDbSettings
    {
        IFlatsSettings FlatsSettings { get; set; }
        IHousesSettings HousesSettings { get; set; }
    }

    public interface IHousesSettings : IWorksheetSettings, ITableSettings
    {
        IRange AddressColumnRange { get; set; }
        bool NonDuplicate { get; set; }
    }

    public interface IFlatsSettings : IWorksheetSettings, ITableSettings
    {
        IRange AddressColumnRange { get; set; }
    }

    public interface ITableSettings
    {
        int TableBeginRow { get; set; }
        int StartRow { get; set; }
    }
}
