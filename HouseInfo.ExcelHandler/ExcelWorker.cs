using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace HouseInfo.ExcelHandler
{
    public  class ExcelWorker
    {
        readonly Settings.ICalcSettings _settingsCalc;
        readonly Settings.IDbSettings _settingsDb;

        Excel.Application _excelApp;

        Excel.Workbook _workbookCalc;
        Excel.Worksheet _worksheetCalc;

        Excel.Workbook _workbookBd;     
        Excel.Worksheet _worksheetFlat;
        Excel.Worksheet _worksheetHouse;


        public ExcelWorker(Settings.ICalcSettings settingsCalc, Settings.IDbSettings settingsDb)
        {
            _settingsCalc = settingsCalc;
            _settingsDb = settingsDb;
        }

        public  void Open(string pathCalc, string pathBd)
        {
            _excelApp = new Excel.Application {Visible = false};
            _workbookCalc = _excelApp.Workbooks.Open(pathCalc);
            _workbookBd = _excelApp.Workbooks.Open(pathBd);

            _worksheetCalc = (Excel.Worksheet) _workbookCalc.Sheets.Item[_settingsCalc.SheetName];           
            _worksheetFlat = (Excel.Worksheet) _workbookBd.Sheets.Item[_settingsDb.FlatsSettings.SheetName];
            _worksheetHouse = (Excel.Worksheet) _workbookBd.Sheets.Item[_settingsDb.HousesSettings.SheetName];            
        }


        public void CopyAddressToCalcSheet(int currentRow)
        {
            _worksheetCalc.Range[_settingsCalc.AddressRange.Begin, _settingsCalc.AddressRange.End].Value 
                = _worksheetFlat.Range[_settingsDb.FlatsSettings.AddressColumnRange.Begin + currentRow, 
                                       _settingsDb.FlatsSettings.AddressColumnRange.End + currentRow].Value;
        }

    

        public void CopyInfoToFlatSheet(int currentRow) 
        {
            foreach (var matchingRange in _settingsCalc.СopiedRangesToFlats)
            {
                    _worksheetFlat.Range[matchingRange.ToRange.Begin + currentRow, matchingRange.ToRange.End + currentRow].Value2 =
                        _worksheetCalc.Range[matchingRange.FromRange.Begin, matchingRange.FromRange.End].Value2;
            }
            //_worksheetFlat.Range[_settingsIdres.sheetFlat.collumn_1 + currentRow, _settingsIdres.sheetFlat.collumn_2 + currentRow].Value = _worksheetCalc.Range[_settings109.f1, _settings109.f2].Value;
            //_worksheetFlat.Range[_settingsIdres.sheetFlat.collumn_3 + currentRow, _settingsIdres.sheetFlat.collumn_17 + currentRow].Value = _worksheetCalc.Range[_settings109.f3, _settings109.f17].Value;
            //_worksheetFlat.Range[_settingsIdres.sheetFlat.collumn_18 + currentRow, _settingsIdres.sheetFlat.collumn_33 + currentRow].Value = _worksheetCalc.Range[_settings109.f18, _settings109.f33].Value;
            //_worksheetFlat.Range[_settingsIdres.sheetFlat.collumn_38 + currentRow].Value = _worksheetCalc.Range[_settings109.f38].Value;
            //_worksheetFlat.Range["F" + currentRow, "AM" + currentRow].Value = _worksheetCalc.Range["F76", "AM76"].Value;

        }

        // нужно для проверки уникальности адресов
        public string AddressOfHouse(int currentRow = 0)
        {
            Excel.Range addressRange;
            string address;
            if (currentRow > 0) //
            {
                addressRange = _worksheetFlat.Range[
                    _settingsDb.FlatsSettings.AddressColumnRange.Begin + currentRow,
                    _settingsDb.FlatsSettings.AddressColumnRange.End + currentRow];
                if (addressRange.Count == 1){
                    return addressRange.Value2;
                }

                address = "";
                
                for (int j = 1; j < addressRange.Columns.Count; j++)
                {
                    address = String.Concat(address, addressRange.Value2[1,j]);
                }         
                return address;
            }

            addressRange = _worksheetFlat.Range[_settingsCalc.AddressRange.Begin, _settingsCalc.AddressRange.End];
            if (addressRange.Count == 1){
                return addressRange.Value2;
            }

                
            address = "";
            for (int j = 1; j < addressRange.Columns.Count; j++)
            {
                address = String.Concat(address, addressRange.Value2[1,j]);
            }         
            return address;
        }

        public void CopyInfoToHouseSheet(int currentRow) 
        {
            foreach (var matchingRange in _settingsCalc.СopiedRangesToHouses)
            {
                _worksheetHouse.Range[matchingRange.ToRange.Begin + currentRow, matchingRange.ToRange.End + currentRow].Value2
                    = _worksheetCalc.Range[matchingRange.FromRange.Begin, matchingRange.FromRange.End].Value2;
            }
            
             //_worksheetHouse.Range[_settingsBd.SheetHouse.collumn_34 + currentRow, _settingsBd.SheetHouse.collumn_51 + currentRow].Value = _worksheetCalc.Range[_settingsCalc.f34, _settingsCalc.f51].Value;
        }

        public bool IsEndFlats(int currentRow = 1)
        {
            return String.IsNullOrWhiteSpace(AddressOfHouse(currentRow));
        }

        public bool IsEndHouses(int currentRow)
        {
            Excel.Range currentRange =
                _worksheetHouse.Range[
                    _settingsDb.HousesSettings.AddressColumnRange.Begin+currentRow,
                    _settingsDb.HousesSettings.AddressColumnRange.End+currentRow];

            if (currentRange.Count == 1)
                return string.IsNullOrWhiteSpace(currentRange.Value2);
            string address = "";
            for (int j = 1; j < currentRange.Columns.Count; j++)
            {
                address = string.Concat(address, currentRange.Value2[1, j]);
            }
            return string.IsNullOrWhiteSpace(address);
            //return string.IsNullOrWhiteSpace(
            //                String.Format("{0}{1}{2}{3}",
            //                    _worksheetHouse.Range[_settingsBd.SheetHouse.collumn_39 + currentRow].Value,
            //                    _worksheetHouse.Range[_settingsBd.SheetHouse.collumn_40 + currentRow].Value,
            //                    _worksheetHouse.Range[_settingsBd.SheetHouse.collumn_41 + currentRow].Value,
            //                    _worksheetHouse.Range[_settingsBd.SheetHouse.collumn_42 + currentRow].Value));
        }
        public void SetVisable(bool isVisable)
        {
            _excelApp.Visible = isVisable;
        }
    }
}
