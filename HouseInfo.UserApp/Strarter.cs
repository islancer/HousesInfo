using System;
using System.Collections.Generic;
using System.Linq;
using HouseInfo.ExcelHandler.Settings;
using HouseInfo.ExcelHandler.Model.Abstract;
using HouseInfo.ExcelHandler.Model;
using HouseInfo.ExcelHandler;
using System.Text.RegularExpressions;

namespace HouseInfo.UserApp
{
    public class Strarter
    {
        public string PathCalc { get; set; }     
        public string PathBd { get; set; }

        private ICalcSettings _settingsCalc;
        private IDbSettings _settingsDb;
        public event Action<string> ShowMessage;

        public void OpenSettings()
        {           
            Regex rgxWhiteSpace = new Regex(@"\s+");
            Regex rgxRangeCells = new Regex(@"^(([a-z]+[1-9][0-9]*-[a-z]+[1-9][0-9]*)|([a-z]+[1-9][0-9]*))$", RegexOptions.IgnoreCase);
            Regex rgxRangeColumns = new Regex(@"^(([a-z]+-[a-z]+)|([a-z]+))$", RegexOptions.IgnoreCase);
            Regex rgxMatchingRange = new Regex(@"^((([a-z]+[1-9][0-9]*-[a-z]+[1-9][0-9]*)|([a-z]+[1-9][0-9]*))=(([a-z]+-[a-z]+)|([a-z]+)))$", RegexOptions.IgnoreCase);

            string message = "";

            #region settingsCalc

            _settingsCalc = new SettingsCalc { SheetName = ExcelSettings.Default.WorksheetNameInCalc};
            if (rgxRangeCells.IsMatch(rgxWhiteSpace.Replace(ExcelSettings.Default.AddressInCalc, "").TrimEnd(';')))
            {
                var splitingAddressInCalc =
                    rgxWhiteSpace.Replace(ExcelSettings.Default.AddressInCalc, "").TrimEnd(';').Split('-');
                if (splitingAddressInCalc.Length == 2)
                {
                    _settingsCalc.AddressRange = new Range
                    {
                        Begin = splitingAddressInCalc[0],
                        End = splitingAddressInCalc[1]
                    };
                }
                else
                    message = string.Format("{0}\r\nОшибка в AddressInCalc: [{1}]", message,
                            ExcelSettings.Default.AddressInCalc);
            }
            else
                message = string.Format("{0}\r\nОшибка в AddressInCalc: [{1}]", message,
                        ExcelSettings.Default.AddressInCalc);
            

            var matchingRangesFlats =
                rgxWhiteSpace.Replace(ExcelSettings.Default.СopiedIntervalsToFlatsWorksheet, "").TrimEnd(';').Split(';');
            if (matchingRangesFlats.Any())
            {
                _settingsCalc.СopiedRangesToFlats = new List<IMatchingRange>();
                foreach (var matchingRange in matchingRangesFlats)
                {
                    if (rgxMatchingRange.IsMatch(matchingRange))
                    {                      
                        var splitMatchingRange = matchingRange.Split('=');

                        
                        var splitFromRange = splitMatchingRange[0].Split('-');
                        Range fromRange = (splitFromRange.Length == 2)
                            ? new Range(splitFromRange[0], splitFromRange[1])
                            : new Range(splitFromRange[0]);
                        
                        var splitToRange = splitMatchingRange[1].Split('-');
                        Range toRange = (splitToRange.Length == 2)
                            ? new Range(splitToRange[0], splitToRange[1])
                            : new Range(splitToRange[0]);
                        _settingsCalc.СopiedRangesToFlats.Add(new MatchingRange(fromRange, toRange));
                    }
                    else
                    {
                        message = string.Format("{0}\r\nОшибка в СopiedIntervalsToFlatsWorksheet: [{1}]", message,
                            matchingRange);
                    }
                }
            }
            else
            {
                message = string.Format("{0}\r\nОшибка в СopiedIntervalsToFlatsWorksheet: [{1}]", message,
                    ExcelSettings.Default.СopiedIntervalsToFlatsWorksheet);
            }


            var matchingRangesHouses =
              rgxWhiteSpace.Replace(ExcelSettings.Default.СopiedIntervalsToHousesWorksheet, "").TrimEnd(';').Split(';');
            if (matchingRangesFlats.Any())
            {
                _settingsCalc.СopiedRangesToHouses = new List<IMatchingRange>();
                foreach (var matchingRange in matchingRangesHouses)
                {
                    if (rgxMatchingRange.IsMatch(matchingRange))
                    {
                        var splitMatchingRange = matchingRange.Split('=');


                        var splitFromRange = splitMatchingRange[0].Split('-');
                        Range fromRange = (splitFromRange.Length == 2)
                            ? new Range(splitFromRange[0], splitFromRange[1])
                            : new Range(splitFromRange[0]);

                        var splitToRange = splitMatchingRange[1].Split('-');
                        Range toRange = (splitToRange.Length == 2)
                            ? new Range(splitToRange[0], splitToRange[1])
                            : new Range(splitToRange[0]);
                        _settingsCalc.СopiedRangesToHouses.Add(new MatchingRange(fromRange, toRange));
                    }
                    else
                    {
                        message = string.Format("{0}\r\nОшибка в СopiedIntervalsToHousesWorksheet: [{1}]", message,
                            matchingRange);
                    }
                }
            }
            else
            {
                message = string.Format("{0}\r\nОшибка в СopiedIntervalsToHousesWorksheet: [{1}]", message,
                    ExcelSettings.Default.СopiedIntervalsToFlatsWorksheet);
            }
            #endregion //settingsCalc

            _settingsDb = new SettingsDb {FlatsSettings = new SettingsFlats(), HousesSettings = new SettingsHouses()};


            _settingsDb.FlatsSettings.SheetName = ExcelSettings.Default.WorksheetFlatsNameInBD;
            _settingsDb.FlatsSettings.TableBeginRow = ExcelSettings.Default.FlatsBeginTableRow;
            _settingsDb.FlatsSettings.StartRow = ExcelSettings.Default.FlatsStartRow;
            if (rgxRangeColumns.IsMatch(rgxWhiteSpace.Replace(ExcelSettings.Default.AddressColumnsInFlats, "").TrimEnd(';')))
            {
                var splitingAddressInFlats =
                    rgxWhiteSpace.Replace(ExcelSettings.Default.AddressColumnsInFlats, "").TrimEnd(';').Split('-');
                if (splitingAddressInFlats.Length == 2)
                {
                    _settingsDb.FlatsSettings.AddressColumnRange = new Range
                    {
                        Begin = splitingAddressInFlats[0],
                        End = splitingAddressInFlats[1]
                    };
                }
                else
                    message = string.Format("{0}\r\nОшибка в AddressColumnsInFlats: [{1}]", message,
                            ExcelSettings.Default.AddressColumnsInFlats);
            }
            else
                message = string.Format("{0}\r\nОшибка в AddressColumnsInFlats: [{1}]", message,
                        ExcelSettings.Default.AddressColumnsInFlats);


            _settingsDb.HousesSettings.SheetName = ExcelSettings.Default.WorksheetHousesNameInBD;
            _settingsDb.HousesSettings.TableBeginRow = ExcelSettings.Default.HousesBeginRow;
            _settingsDb.HousesSettings.StartRow = ExcelSettings.Default.HousesStartRow;
            _settingsDb.HousesSettings.NonDuplicate = ExcelSettings.Default.NonDublicateHouses;
            if (rgxRangeColumns.IsMatch(rgxWhiteSpace.Replace(ExcelSettings.Default.AddressColumnsInHouses, "").TrimEnd(';')))
            {
                var splitingAddressInHouses =
                    rgxWhiteSpace.Replace(ExcelSettings.Default.AddressColumnsInHouses, "").TrimEnd(';').Split('-');
                if (splitingAddressInHouses.Length == 2)
                {
                    _settingsDb.HousesSettings.AddressColumnRange = new Range
                    {
                        Begin = splitingAddressInHouses[0],
                        End = splitingAddressInHouses[1]
                    };
                }
                else
                    message = string.Format("{0}\r\nОшибка в AddressColumnsInHouses: [{1}]", message,
                            ExcelSettings.Default.AddressColumnsInHouses);
            }
            else
                message = string.Format("{0}\r\nОшибка в AddressColumnsInHouses: [{1}]", message,
                        ExcelSettings.Default.AddressColumnsInHouses);

            if ((ShowMessage != null) && !String.IsNullOrEmpty(message)) ShowMessage(message);
        }
        public void Start()
        {
            string message = null;
            if (!String.IsNullOrWhiteSpace(PathCalc) && !String.IsNullOrWhiteSpace(PathBd))
            {
                var worker = new Worker(PathCalc, _settingsCalc, PathBd, _settingsDb);
                
                try
                {
                    if (ShowMessage != null) worker.StatusMassage += ShowMessage;               
                    worker.Work();
                }
                catch (Exception ex)
                {
                    message = ex.Message;

                }
                try
                {
                    worker.ExcelWorker.SetVisable(true);
                }
                catch (Exception ex)
                {
                    message = (message != null) ? (message + "\r\n" + ex.Message) : ex.Message;

                }
                
            }
            else
            {
                message = "!    Не указан путь к файлам   или указан неверно   !";
            }
            
            if ((ShowMessage != null) && (message != null)) ShowMessage(message);
        }

    }
}
