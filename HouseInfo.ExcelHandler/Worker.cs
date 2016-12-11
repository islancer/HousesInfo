using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HouseInfo.ExcelHandler
{
    public class Worker
    {
        public ExcelWorker excelWorker;
        private bool _cancelled = false;
        public string path109;
        public string pathIDRES;
        public HouseInfo.ExcelHandler.Settings.Settings_109 settings_109;
        public HouseInfo.ExcelHandler.Settings.Settings_IDRES settings_IDRES;
        
       // public Worker() { }
        public Worker(string path109 ,HouseInfo.ExcelHandler.Settings.Settings_109 settings_109, string pathIDRES, HouseInfo.ExcelHandler.Settings.Settings_IDRES settings_IDRES)
        {
            _cancelled = false;
            this.path109 = path109;
            this.pathIDRES = pathIDRES; 
            this.settings_109 = settings_109;
            this.settings_IDRES = settings_IDRES;
            
        }
        public void Cancel()
        {
            _cancelled = true;
        }

        public bool Work()
        {
            excelWorker = new ExcelWorker(settings_109, settings_IDRES);
            excelWorker.Open(path109, pathIDRES);
            IDictionary<string, bool> addressHouses = new Dictionary<string, bool>();
            
            int rowFlat = (settings_IDRES.sheetFlat.startRow > 0) ? settings_IDRES.sheetFlat.startRow : Math.Max(settings_IDRES.sheetFlat.tableStartRow, 1);
            int rowHouse = FindStartRowHousesSheet();

            while(!excelWorker.isEndFlats(rowFlat) && !_cancelled)
            {
                excelWorker.CopyAddressToCalcSheet(rowFlat);
                excelWorker.CopyInfoToFlatSheet(rowFlat);
                if (!addressHouses.ContainsKey(excelWorker.addressOfHouse(rowFlat)))
                {
                    addressHouses.Add(excelWorker.addressOfHouse(rowFlat), true);
                    excelWorker.CopyInfoToHouseSheet(rowHouse);
                    rowHouse++;
                }
                if (rowFlat % 5 == 0)
                    StatusMassage(String.Format("row {0}", rowFlat));
                rowFlat++;

            }
            StatusMassage(String.Format("finish row {0}", rowFlat));
            return _cancelled;
        }

        private int FindStartRowHousesSheet()
        {
            StatusMassage("Поиск места для заполнения в листе с онформацией о домах.");
            int row = settings_IDRES.sheetHouse.startRow;
            if (row < settings_IDRES.sheetHouse.tableStartRow)
            {
                row = (settings_IDRES.sheetHouse.tableStartRow > 0) ? settings_IDRES.sheetHouse.tableStartRow : 1;
                while (!excelWorker.isEndHouses(row))
                {
                    row++;
                }
            }
            return row;
        }

        public event Action<string> StatusMassage;
        public event Action<bool> WorkCompleted;
    }
}
