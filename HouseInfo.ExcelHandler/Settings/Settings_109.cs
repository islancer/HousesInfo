using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

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
    public  class Settings_109
    {
        public Settings_109() {}

        public Settings_109(bool isSetDefault)
        {
            if (isSetDefault)
                setDefault();
        }

        

        /// <summary>
        /// Название листа для расчётов
        /// </summary>
        public string sheetName { get; set; }

        /// <summary>
        /// тип улицы
        /// </summary>
        public string a { get; set; }

        /// <summary>
        /// название улицы
        /// </summary>
        public string b { get; set; }

        /// <summary>
        /// номер дома
        /// </summary>
        public string c { get; set; }

        /// <summary>
        /// номер корпуса
        /// </summary>
        public string d { get; set; }

        /// <summary>
        /// номер кв.
        /// </summary>
        public string e { get; set; }

        #region ячейки 1-51

        public string f1 { get; set; }
        public string f2 { get; set; }
        public string f3 { get; set; }
        public string f4 { get; set; }
        public string f5 { get; set; }
        public string f6 { get; set; }
        public string f7 { get; set; }
        public string f8 { get; set; }
        public string f9 { get; set; }
        public string f10 { get; set; }
        public string f11 { get; set; }
        public string f12 { get; set; }
        public string f13 { get; set; }
        public string f14 { get; set; }
        public string f15 { get; set; }
        public string f16 { get; set; }
        public string f17 { get; set; }
        public string f18 { get; set; }
        public string f19 { get; set; }
        public string f20 { get; set; }
        public string f21 { get; set; }
        public string f22 { get; set; }
        public string f23 { get; set; }
        public string f24 { get; set; }
        public string f25 { get; set; }
        public string f26 { get; set; }
        public string f27 { get; set; }
        public string f28 { get; set; }
        public string f29 { get; set; }
        public string f30 { get; set; }
        public string f31 { get; set; }
        public string f32 { get; set; }
        public string f33 { get; set; }
        public string f34 { get; set; }
        public string f35 { get; set; }
        public string f36 { get; set; }
        public string f37 { get; set; }
        public string f38 { get; set; }
        public string f39 { get; set; }
        public string f40 { get; set; }
        public string f41 { get; set; }
        public string f42 { get; set; }
        public string f43 { get; set; }
        public string f44 { get; set; }
        public string f45 { get; set; }
        public string f46 { get; set; }
        public string f47 { get; set; }
        public string f48 { get; set; }
        public string f49 { get; set; }
        public string f50 { get; set; }
        public string f51 { get; set; }

        #endregion ячейки f01-f51

        public void setDefault()
        {
            sheetName = "лист запрос";

            this.a = "A32";
            this.b = "B32";
            this.c = "C32";
            this.d = "D32";
            this.e = "E32";

            #region базовые адреса ячеек 1-51

            this.f1 = "A36";
            this.f2 = "B36";
            this.f3 = "C36";
            this.f4 = "D36";
            this.f5 = "E36";
            this.f6 = "F36";
            this.f7 = "G36";
            this.f8 = "H36";
            this.f9 = "I36";
            this.f10 = "J36";
            this.f11 = "K36";
            this.f12 = "L36";
            this.f13 = "M36";
            this.f14 = "N36";
            this.f15 = "O36";
            this.f16 = "P36";
            this.f17 = "Q36";
            this.f18 = "B39";
            this.f19 = "C39";
            this.f20 = "D39";
            this.f21 = "E39";
            this.f22 = "F39";
            this.f23 = "G39";
            this.f24 = "H39";
            this.f25 = "I39";
            this.f26 = "J39";
            this.f27 = "K39";
            this.f28 = "L39";
            this.f29 = "M39";
            this.f30 = "N39";
            this.f31 = "O39";
            this.f32 = "P39";
            this.f33 = "Q39";
            this.f34 = "A43";
            this.f35 = "B43";
            this.f36 = "C43";
            this.f37 = "D43";
            this.f38 = "E43";
            this.f39 = "F43";
            this.f40 = "G43";
            this.f41 = "H43";
            this.f42 = "I43";
            this.f43 = "J43";
            this.f44 = "K43";
            this.f45 = "L43";
            this.f46 = "M43";
            this.f47 = "N43";
            this.f48 = "O43";
            this.f49 = "P43";
            this.f50 = "Q43";
            this.f51 = "R43";

            #endregion базовые адреса ячеек 1-51
        }


        //{
        //    XmlSerializer formatter = new XmlSerializer(typeof(HouseInfo.ExcelHandler.Settings.Settings_109));
        //    using (FileStream fs = new FileStream("HouseInfo.f_109.config", FileMode.Create))
        //    {
        //        formatter.Serialize(fs, this);

        //        Console.WriteLine("Объект сериализован");
        //    }
        //}

    }
    
}
