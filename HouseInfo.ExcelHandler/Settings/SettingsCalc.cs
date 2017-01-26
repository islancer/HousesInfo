using System;
using System.Collections.Generic;
using HouseInfo.ExcelHandler.Model.Abstract;

namespace HouseInfo.ExcelHandler.Settings
{
//-------------------------------
//
//      Для каждой ячейки создано своё св-во.
//      Не использованны массивы для ячеек a-d, f1-f51 для простоты ассоциаций и чтению кода и файлов ностройки.
//
//-------------------------------


    /// <summary>
    /// Адреса всех ячеек вспомогательного файла для расчётов параметров квартиры/дома
    /// </summary>
    [Serializable]
    public  class SettingsCalc : ICalcSettings
    {
        public SettingsCalc() {}
        public string SheetName { get; set; }

        public IRange AddressRange { get; set; }

        public IList<IMatchingRange> СopiedRangesToFlats { get; set; }

        public IList<IMatchingRange> СopiedRangesToHouses { get; set; }
      
    }
    
}
