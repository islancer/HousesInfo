using System;

using System.IO;

using System.Windows.Forms;
using System.Xml.Serialization;

namespace HouseInfo.UserApp
{
    public static class Menu
    {
        private static readonly Strarter StrarterWork;
        static string _pathCalc = "";
        static ExcelHandler.Settings.SettingsCalc _settingsCalc;
        static string _pathBd = "";
        static ExcelHandler.Settings.SettingsDb _settingsDb;
        private static readonly OpenFileDialog OpenFileDialogPath;

        static Menu()
        {
            StrarterWork = new Strarter();
            StrarterWork.ShowMessage += ShowMassage;
            OpenFileDialogPath = new OpenFileDialog();
            
        }

        public static void ShowMenu()
        {

           StrarterWork.OpenSettings();
            
           

            bool exit = false;
            while (!exit)
            {

                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("-------------");
                Console.WriteLine("     Меню");
                Console.WriteLine("-------------");
                Console.WriteLine();
                Console.WriteLine("1 - Путь к файлам");
                Console.WriteLine("2 - Запуск на выполнение");
                
                Console.WriteLine("0 - Выход");

                switch (Console.ReadKey().Key)
                {
                    case ConsoleKey.D1:
                        {
                            Console.Clear();
                            MenuPathFiles();
                            break;
                        }
                    case ConsoleKey.NumPad1:
                        {
                            Console.Clear();
                            MenuPathFiles();
                            break;
                        }
                    case ConsoleKey.D2:
                        {
                            Console.Clear();
                            StrarterWork.Start();
                            break;
                        }
                    case ConsoleKey.NumPad2:
                        {
                            Console.Clear();
                            StrarterWork.Start();
                            break;
                        }
              
                    case ConsoleKey.D0:
                        { exit = true; break; }
                    case ConsoleKey.NumPad0:
                        { exit = true; break; }
                    case ConsoleKey.Escape:
                        { exit = true; break; }
                    default:
                        { Console.Clear(); break; }
                }
            }

        }

        static void MenuPathFiles()
        {

            
            OpenFileDialogPath.RestoreDirectory = true;
            OpenFileDialogPath.Filter = "Excel Files(*.xls;*.xlsx;)|*.xls;*.xlsx;|All files (*.*)|*.*";
            bool exit = false;
            while (!exit)
            {

                Console.WriteLine("Файл для расчётов: {0}", StrarterWork.PathCalc);
                Console.WriteLine("Файл для заполнения: {0}", StrarterWork.PathBd);
                Console.WriteLine();
                Console.WriteLine("<-------------");
                Console.WriteLine(" Путь к файлам");
                Console.WriteLine("<-------------");
                Console.WriteLine();
                Console.WriteLine("1 - Путь к файлу c расчётами");
                Console.WriteLine("2 - Путь к файлу для заполнения");
                Console.WriteLine("0 - Назад;");
                switch (Console.ReadKey().Key)
                {
                    case ConsoleKey.D1:
                        {                          
                            StrarterWork.PathCalc = GetPath("Укажите путь к файлу c расчётами") ?? StrarterWork.PathCalc;
                            Console.WriteLine();
                            break;
                        }
                    case ConsoleKey.NumPad1:
                        {
                            StrarterWork.PathCalc = GetPath("Укажите путь к файлу c расчётами") ?? StrarterWork.PathCalc;
                             Console.WriteLine();
                            break;
                        }
                    case ConsoleKey.D2:
                        {
                            StrarterWork.PathBd = GetPath("Укажите путь к файлу для заполнения") ?? StrarterWork.PathCalc;
                            Console.WriteLine();
                            break;
                        }
                    case ConsoleKey.NumPad2:
                        {
                            StrarterWork.PathBd = GetPath("Укажите путь к файлу для заполнения") ?? StrarterWork.PathCalc;
                            Console.WriteLine();
                            break;
                        }
                    case ConsoleKey.D0:
                        { Console.Clear(); exit = true; break; }
                    case ConsoleKey.NumPad0:
                        { Console.Clear(); exit = true; break; }
                    case ConsoleKey.Escape:
                        { Console.Clear(); exit = true; break; }


                    default:
                        { Console.Clear(); break; }

                }
            }
        }

        static string GetPath(string title)
        {
            OpenFileDialogPath.Title = title;
            if (OpenFileDialogPath.ShowDialog() == DialogResult.OK)
            {
                
                Console.Clear();
                return OpenFileDialogPath.FileName;
            }
            else
            {
                Console.Clear();
                Console.WriteLine("Действие было отменено");

            }
            return null;
        }
        static void ShowMassage(string msg)
        {
            Console.Clear();
            Console.WriteLine(msg);
        }

        //static void SaveSettings(ExcelHandler.Settings.SettingsCalc settingsCalc, ExcelHandler.Settings.SettingsDb settingsBd)
        //{
        //    try
        //    {
        //        var formatter = new XmlSerializer(typeof(ExcelHandler.Settings.SettingsCalc));
        //        using (FileStream fs = new FileStream("HouseInfo.f_109.config", FileMode.Create))
        //        {
        //            formatter.Serialize(fs, settingsCalc);
        //        }


        //        XmlSerializer formatter2 = new XmlSerializer(typeof(ExcelHandler.Settings.SettingsDb));
        //        using (FileStream fs = new FileStream("HouseInfo.f_IDRES.config", FileMode.Create))
        //        {
        //            formatter2.Serialize(fs, settingsBd);
        //        }
        //        Console.WriteLine("Файлы настроек созданы.");
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.Message);
        //    }
        //}

      //  static void OpenSettings(out ExcelHandler.Settings.SettingsCalc settingsCalc, out ExcelHandler.Settings.SettingsDb settingsDb)
      //  {
            
            //try
            //{
            //    XmlSerializer formatter = new XmlSerializer(typeof(ExcelHandler.Settings.SettingsCalc));
            //    using (FileStream fs = new FileStream("HouseInfo.f_109.config", FileMode.Open))
            //    {
            //        settingsCalc = (ExcelHandler.Settings.SettingsCalc)formatter.Deserialize(fs);
            //    }

            //    Console.WriteLine("Настройки из 'HouseInfo.f_109.config' загружены.");
            //}
            //catch (Exception ex)
            //{
            //    settingsCalc = new ExcelHandler.Settings.SettingsCalc(true);
            //    Console.WriteLine(ex.Message);
            //    Console.WriteLine("Использованы базовые натройки.");
            //}
            //try
            //{
            //    XmlSerializer formatter = new XmlSerializer(typeof(ExcelHandler.Settings.SettingsDb));
            //    using (FileStream fs = new FileStream("HouseInfo.f_IDRES.config", FileMode.Open))
            //    {
            //        settingsDb = (ExcelHandler.Settings.SettingsDb)formatter.Deserialize(fs);
            //    }

            //    Console.WriteLine("Настройки из 'HouseInfo.f_IDRES.config' загружены.");
            //}
            //catch (Exception ex)
            //{
            //    settingsDb = new ExcelHandler.Settings.SettingsDb();
            //    Console.WriteLine(ex.Message);
            //    Console.WriteLine("Использованы базовые натройки.");
            //}

            
      //  }

      
    }
}
