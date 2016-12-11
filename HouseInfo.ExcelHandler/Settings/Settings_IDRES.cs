using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HouseInfo.ExcelHandler.Settings
{
    [Serializable]
    public class Settings_IDRES
    {
        public Settings_IDRES() {
            sheetFlat = new SheetFlat();
            sheetHouse = new SheetHouses();
        }

        public Settings_IDRES(bool isSetDefault)
        {
            if (isSetDefault)
            {
                sheetFlat = new SheetFlat(isSetDefault);
                sheetHouse = new SheetHouses(isSetDefault);
            }
        }

        public SheetFlat sheetFlat { get; set; }
        public SheetHouses sheetHouse { get; set; }

        [Serializable]
        public class SheetFlat
        {
            public SheetFlat() { }

            public SheetFlat(bool isSetDefault) 
            {
                if (isSetDefault)
                    setDefault();
            }

            /// <summary>
            /// Название листа
            /// </summary>
            public string sheetName { get; set; }

            /// <summary>
            /// Строка с которой начинается таблица
            /// </summary>
            public int tableStartRow { get; set; }

            /// <summary>
            /// Номер строки с которой начинаем доставать/сохранять данные
            /// </summary>
            public int startRow { get; set; }

            /// <summary>
            /// тип улицы
            /// </summary>
            public string collumn_a { get; set; }

            /// <summary>
            /// название улицы
            /// </summary>
            public string collumn_b { get; set; }

            /// <summary>
            /// номер дома
            /// </summary>
            public string collumn_c { get; set; }

            /// <summary>
            /// номер корпуса
            /// </summary>
            public string collumn_d { get; set; }

            /// <summary>
            /// номер кв.
            /// </summary>

            public string collumn_e { get; set; }

            #region ячейки 01-33, 38
           
            public string collumn_1 { get; set; }
            public string collumn_2 { get; set; }
            public string collumn_3 { get; set; }
            public string collumn_4 { get; set; }
            public string collumn_5 { get; set; }
            public string collumn_6 { get; set; }
            public string collumn_7 { get; set; }
            public string collumn_8 { get; set; }
            public string collumn_9 { get; set; }
            public string collumn_10 { get; set; }
            public string collumn_11 { get; set; }
            public string collumn_12 { get; set; }
            public string collumn_13 { get; set; }
            public string collumn_14 { get; set; }
            public string collumn_15 { get; set; }
            public string collumn_16 { get; set; }
            public string collumn_17 { get; set; }
            public string collumn_18 { get; set; }
            public string collumn_19 { get; set; }
            public string collumn_20 { get; set; }
            public string collumn_21 { get; set; }
            public string collumn_22 { get; set; }
            public string collumn_23 { get; set; }
            public string collumn_24 { get; set; }
            public string collumn_25 { get; set; }
            public string collumn_26 { get; set; }
            public string collumn_27 { get; set; }
            public string collumn_28 { get; set; }
            public string collumn_29 { get; set; }
            public string collumn_30 { get; set; }
            public string collumn_31 { get; set; }
            public string collumn_32 { get; set; }
            public string collumn_33 { get; set; }
            public string collumn_38 { get; set; }
            
            #endregion ячейки 1-33, 38

            public void setDefault()
            {
                this.sheetName = "tab1";
                this.tableStartRow = 5;
                this.startRow = 5;
                this.collumn_a = "B";
                this.collumn_b = "C";
                this.collumn_c = "D";
                this.collumn_d = "E";
                this.collumn_e = "F";

                #region базовые настройки для столбцов 1-33, 38
                this.collumn_1 = "G";
                this.collumn_2 = "H";
                this.collumn_3 = "N";
                this.collumn_4 = "O";
                this.collumn_5 = "P";
                this.collumn_6 = "Q";
                this.collumn_7 = "R";
                this.collumn_8 = "S";
                this.collumn_9 = "T";
                this.collumn_10 = "U";
                this.collumn_11 = "V";
                this.collumn_12 = "W";
                this.collumn_13 = "X";
                this.collumn_14 = "Y";
                this.collumn_15 = "Z";
                this.collumn_16 = "AA";
                this.collumn_17 = "AB";
                this.collumn_18 = "AC";
                this.collumn_19 = "AD";
                this.collumn_20 = "AE";
                this.collumn_21 = "AF";
                this.collumn_22 = "AG";
                this.collumn_23 = "AH";
                this.collumn_24 = "AI";
                this.collumn_25 = "AJ";
                this.collumn_26 = "AK";
                this.collumn_27 = "AL";
                this.collumn_28 = "AM";
                this.collumn_29 = "AN";
                this.collumn_30 = "AO";
                this.collumn_31 = "AP";
                this.collumn_32 = "AQ";
                this.collumn_33 = "AR";
                this.collumn_38 = "A";

                #endregion
            }
        }

        [Serializable]
        public class SheetHouses
        {

             public SheetHouses() { }

             public SheetHouses(bool isSetDefault) 
            {
                if (isSetDefault)
                    setDefault();
            }

             

            public string sheetName { get; set; }

            /// <summary>
            /// Строка с которой начинается таблица
            /// </summary>
            public int tableStartRow { get; set; }

            public int startRow { get; set; }

            public bool nonDuplicate { get; set; }

            #region ячейки 34-51

            public string collumn_34 { get; set; }
            public string collumn_35 { get; set; }
            public string collumn_36 { get; set; }
            public string collumn_37 { get; set; }
            public string collumn_38 { get; set; }
            public string collumn_39 { get; set; }
            public string collumn_40 { get; set; }
            public string collumn_41 { get; set; }
            public string collumn_42 { get; set; }
            public string collumn_43 { get; set; }
            public string collumn_44 { get; set; }
            public string collumn_45 { get; set; }
            public string collumn_46 { get; set; }
            public string collumn_47 { get; set; }
            public string collumn_48 { get; set; }
            public string collumn_49 { get; set; }
            public string collumn_50 { get; set; }
            public string collumn_51 { get; set; }

            #endregion

            public void setDefault()
            {
                this.sheetName = "tab2";
                this.tableStartRow = 5;
                this.startRow = 0;
                this.nonDuplicate = true;
                #region базовые настройки для столбцов 33-51

                this.collumn_34 = "A";
                this.collumn_35 = "B";
                this.collumn_36 = "C";
                this.collumn_37 = "D";
                this.collumn_38 = "E";
                this.collumn_39 = "F";
                this.collumn_40 = "G";
                this.collumn_41 = "H";
                this.collumn_42 = "I";
                this.collumn_43 = "J";
                this.collumn_44 = "K";
                this.collumn_45 = "L";
                this.collumn_46 = "M";
                this.collumn_47 = "N";
                this.collumn_48 = "O";
                this.collumn_49 = "P";
                this.collumn_50 = "Q";
                this.collumn_51 = "R";

                #endregion базовые настройки для столбцов 33-51
            }
        }
    }
}
