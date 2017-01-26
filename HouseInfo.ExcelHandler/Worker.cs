using System;
using System.Collections.Generic;
using HouseInfo.ExcelHandler.Settings;

namespace HouseInfo.ExcelHandler
{
    public class Worker
    {
        public ExcelWorker ExcelWorker { get; set; }       
        public string PathCalc { get; set; }
        public string PathBd { get; set; }
        public ICalcSettings SettingsCalc { get; set; }
        public IDbSettings SettingsDb { get; set; }

        // public Worker() { }
        public Worker(string pathCalc, ICalcSettings settingsCalc, string pathBd, IDbSettings settingsDb)
        {
           
            PathCalc = pathCalc;
            PathBd = pathBd;
            SettingsCalc = settingsCalc;
            SettingsDb = settingsDb;
            
        }
       
        public void Work()
        {
            ExcelWorker = new ExcelWorker(SettingsCalc, SettingsDb);
            ExcelWorker.Open(PathCalc, PathBd);
            IDictionary<string, bool> addressHouses = new Dictionary<string, bool>();

            int rowFlat = (SettingsDb.FlatsSettings.StartRow > 0) ? SettingsDb.FlatsSettings.StartRow : Math.Max(SettingsDb.FlatsSettings.TableBeginRow, 1);
            int rowHouse = FindStartRowHousesSheet();

            ITimeTest timeTest = new TimeTotalTest();
            int iter = 0;
            timeTest.Start(iter);
            while (!ExcelWorker.IsEndFlats(rowFlat))
            {
                
                ExcelWorker.CopyAddressToCalcSheet(rowFlat);
                ExcelWorker.CopyInfoToFlatSheet(rowFlat);
                if (SettingsDb.HousesSettings.NonDuplicate)
                {
                    if (!addressHouses.ContainsKey(ExcelWorker.AddressOfHouse(rowFlat)))
                    {
                        addressHouses.Add(ExcelWorker.AddressOfHouse(rowFlat), true);
                        ExcelWorker.CopyInfoToHouseSheet(rowHouse);
                        rowHouse++;
                    }
                }
                else
                {
                    ExcelWorker.CopyInfoToHouseSheet(rowHouse);
                    rowHouse++;
                }
                if (rowFlat % 100 == 0)
                    if (StatusMassage != null) StatusMassage(String.Format("row {0}", rowFlat));
                rowFlat++;

                
                iter++;
                //if (iter >= 100) break;
            }
            timeTest.Stop(iter);
            timeTest.Save("address.txt");

            if (StatusMassage != null) StatusMassage(String.Format("finish row {0}", rowFlat));
        }


        private int FindStartRowHousesSheet()
        {
            if (StatusMassage != null) StatusMassage("Поиск места для заполнения в листе с информацией о домах.");
            int row = SettingsDb.HousesSettings.StartRow;
            if (row < SettingsDb.HousesSettings.TableBeginRow)
            {
                row = (SettingsDb.HousesSettings.TableBeginRow > 0) ? SettingsDb.HousesSettings.TableBeginRow : 1;
                while (!ExcelWorker.IsEndHouses(row))
                {
                    row++;
                }
            }
            return row;
        }

        public event Action<string> StatusMassage;
        
    }
}
