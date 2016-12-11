using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace HouseInfo.ExcelHandler
{
    public  class ExcelWorker
    {
        HouseInfo.ExcelHandler.Settings.Settings_109 settings_109;
        HouseInfo.ExcelHandler.Settings.Settings_IDRES settings_IDRES;

        Excel.Application excelApp;
        Excel.Workbook workbook109;
        Excel.Workbook workbookIDRES;
        Excel.Worksheet worksheetCalc;
        Excel.Worksheet worksheetFlat;
        Excel.Worksheet worksheetHouse;

        public ExcelWorker() 
        {
            settings_109 = new HouseInfo.ExcelHandler.Settings.Settings_109(true);
            settings_IDRES = new HouseInfo.ExcelHandler.Settings.Settings_IDRES(true);
        }

        public ExcelWorker(HouseInfo.ExcelHandler.Settings.Settings_109 settings_109, HouseInfo.ExcelHandler.Settings.Settings_IDRES settings_IDRES)
        {
            this.settings_109 = settings_109;
            this.settings_IDRES = settings_IDRES;
        }

        public  void Open(string path109, string pathIDRES)
        {
            excelApp = new Excel.Application();
            excelApp.Visible = false;
            workbook109 = excelApp.Workbooks.Open(path109);
            workbookIDRES = excelApp.Workbooks.Open(pathIDRES);

            worksheetCalc = workbook109.Sheets.get_Item(settings_109.sheetName);           
            worksheetFlat = workbookIDRES.Sheets.get_Item(settings_IDRES.sheetFlat.sheetName);
            worksheetHouse = workbookIDRES.Sheets.get_Item(settings_IDRES.sheetHouse.sheetName);            
        }

        public void CopyAddressToCalcSheet(int currentRow)
        {
            worksheetCalc.get_Range(settings_109.a).Value = worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_a + currentRow.ToString()).Value;
            worksheetCalc.get_Range(settings_109.b).Value = worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_b + currentRow.ToString()).Value;
            worksheetCalc.get_Range(settings_109.c).Value = worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_c + currentRow.ToString()).Value;
            worksheetCalc.get_Range(settings_109.d).Value = worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_d + currentRow.ToString()).Value;
            worksheetCalc.get_Range(settings_109.e).Value = worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_e + currentRow.ToString()).Value;
        }

        public void CopyInfoToFlatSheet(int currentRow) // быстрее чем через рефлексию
        {
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_1 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f1).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_2 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f2).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_3 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f3).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_4 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f4).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_5 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f5).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_6 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f6).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_7 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f7).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_8 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f8).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_9 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f9).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_10 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f10).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_11 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f11).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_12 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f12).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_13 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f13).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_14 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f14).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_15 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f15).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_16 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f16).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_17 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f17).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_18 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f18).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_19 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f19).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_20 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f20).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_21 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f21).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_22 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f22).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_23 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f23).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_24 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f24).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_25 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f25).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_26 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f26).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_27 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f27).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_28 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f28).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_29 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f29).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_30 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f30).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_31 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f31).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_32 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f32).Value;
            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_33 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f33).Value;

            worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_38 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f38).Value;
        }

        // нужно для проверки уникальности адресов
        public string addressOfHouse(int currentRow = 0)
        {
            if (currentRow > 0)
                return String.Format("{0} {1} {2} {3}", worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_a + currentRow.ToString()).Value,
                                                        worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_b + currentRow.ToString()).Value,
                                                        worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_c + currentRow.ToString()).Value,
                                                        worksheetFlat.get_Range(settings_IDRES.sheetFlat.collumn_d + currentRow.ToString()).Value
                                    );
            return String.Format("{0} {1} {2} {3}", worksheetCalc.get_Range(settings_109.a).Value,
                                                    worksheetCalc.get_Range(settings_109.b).Value,
                                                    worksheetCalc.get_Range(settings_109.c).Value,
                                                    worksheetCalc.get_Range(settings_109.d).Value
                                );
        }

        public void CopyInfoToHouseSheet(int currentRow) // опять быстрее чем через рефлексию
        {
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_34 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f34).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_35 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f35).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_36 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f36).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_37 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f37).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_38 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f38).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_39 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f39).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_40 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f40).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_41 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f41).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_42 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f42).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_43 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f43).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_44 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f44).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_45 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f45).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_46 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f46).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_47 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f47).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_48 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f48).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_49 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f49).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_50 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f50).Value;
            worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_51 + currentRow.ToString()).Value = worksheetCalc.get_Range(settings_109.f51).Value;
        }

        public bool isEndFlats(int currentRow = 0)
        {
            return String.IsNullOrWhiteSpace(addressOfHouse(currentRow));
        }

        public bool isEndHouses(int currentRow)
        {

            return String.IsNullOrWhiteSpace(
                            String.Format("{0}{1}{2}{3}",
                                worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_39 + currentRow.ToString()).Value,
                                worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_40 + currentRow.ToString()).Value,
                                worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_41 + currentRow.ToString()).Value,
                                worksheetHouse.get_Range(settings_IDRES.sheetHouse.collumn_42 + currentRow.ToString()).Value));
        }
        public void SetVisable(bool isVisable)
        {
            excelApp.Visible = isVisable;
        }
    }
}
