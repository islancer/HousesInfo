using HouseInfo.ExcelHandler.Model.Abstract;

namespace HouseInfo.ExcelHandler.Settings
{

    public class SettingsDb : IDbSettings
    {
        public IFlatsSettings FlatsSettings { get; set; }
        public IHousesSettings HousesSettings { get; set; }
    }

    public class SettingsFlats : IFlatsSettings
    {
        public string SheetName { get; set; }
        public int TableBeginRow { get; set; }
        public int StartRow { get; set; }
        public IRange AddressColumnRange { get; set; }
    }

    public class SettingsHouses : IHousesSettings
    {
        public string SheetName { get; set; }
        public int TableBeginRow { get; set; }
        public int StartRow { get; set; }
        public IRange AddressColumnRange { get; set; }
        public bool NonDuplicate { get; set; }
    }

   
    
}
