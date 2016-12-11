using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Windows.Forms;

namespace HouseInfo.UserApp
{
    class Program
    {
        static string path109 = "";
        static ExcelHandler.Settings.Settings_109 settings_109;
        static string pathIDRES = "";
        static ExcelHandler.Settings.Settings_IDRES settings_IDRES;




        [STAThreadAttribute]
        static void Main(string[] args)
        {
            Menu();
        }




        static void Menu()
        {
            
            
            OpenSettings(out settings_109, out settings_IDRES);

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
                Console.WriteLine("7 - Создать файлы с настройками");
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
                            Start();
                            break;
                        }
                    case ConsoleKey.NumPad2:
                        {
                            Console.Clear();
                            Start();
                            break;
                        }

                    case ConsoleKey.D7:
                        { 
                            Console.Clear();
                            SaveSettings(settings_109, settings_IDRES);
                            break; }
                    case ConsoleKey.NumPad7:
                        {
                            Console.Clear();
                            SaveSettings(settings_109, settings_IDRES);
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
            
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Filter = "Excel Files(*.xls;*.xlsx;)|*.xls;*.xlsx;|All files (*.*)|*.*";
            bool exit = false;
            while (!exit)
            {
                
                Console.WriteLine("Файл для расчётов: {0}", path109);
                Console.WriteLine("Файл для заполнения: {0}", pathIDRES);
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
                            openFileDialog.Title = "Укажите путь к файлу c расчётами";
                            if (openFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                path109 = openFileDialog.FileName;
                                Console.Clear();
                            }
                            else
                            {
                                Console.Clear();
                                Console.WriteLine("Действие было отменено");
                                
                            }
                            Console.WriteLine();
                            break;
                        }
                    case ConsoleKey.NumPad1:
                        {
                            openFileDialog.Title = "Укажите путь к файлу c расчётами";
                            if (openFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                path109 = openFileDialog.FileName;
                                Console.Clear();
                            }
                            else
                            {
                                Console.Clear();
                                Console.WriteLine("Действие было отменено");

                            }
                            Console.WriteLine();
                            break;
                        }
                    case ConsoleKey.D2:
                        {
                            openFileDialog.Title = "Укажите путь к файлу для заполнения";
                            if (openFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                pathIDRES = openFileDialog.FileName;
                                Console.Clear();
                            }
                            else
                            {
                                Console.Clear();
                                Console.WriteLine("Действие было отменено");
                                
                            }
                            Console.WriteLine();
                            break;
                        }
                    case ConsoleKey.NumPad2:
                        {
                            openFileDialog.Title = "Укажите путь к файлу для заполнения";
                            if (openFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                pathIDRES = openFileDialog.FileName;
                                Console.Clear();
                            }
                            else
                            {
                                Console.Clear();
                                Console.WriteLine("Действие было отменено");

                            }
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
                        { Console.Clear();  break; }
                    
                }
            }
        }
        public static void Start()
        {
            if (!String.IsNullOrWhiteSpace(path109) && !String.IsNullOrWhiteSpace(pathIDRES))
            {
                HouseInfo.ExcelHandler.Worker  worker = new ExcelHandler.Worker(path109, settings_109, pathIDRES, settings_IDRES);;
                try
                {                                      
                    worker.StatusMassage += ShowMassage;
                    
                    worker.Work();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    
                }
                try
                {
                    worker.excelWorker.SetVisable(true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                }    
                    
            }
            else
            {
                Console.WriteLine("!    Не указан путь к файлам   или указан неверно   !");
            }
            //Console.WriteLine("Для продолжения нажмите любую клавишу.");
            //Console.ReadKey();
        }

        static void SaveSettings(ExcelHandler.Settings.Settings_109 settings_109, ExcelHandler.Settings.Settings_IDRES settings_IDRES)
        {
            try
            {              
                XmlSerializer formatter = new XmlSerializer(typeof(HouseInfo.ExcelHandler.Settings.Settings_109));
                using (FileStream fs = new FileStream("HouseInfo.f_109.config", FileMode.Create))
                {
                    formatter.Serialize(fs, settings_109);
                }


                XmlSerializer formatter2 = new XmlSerializer(typeof(HouseInfo.ExcelHandler.Settings.Settings_IDRES));
                using (FileStream fs = new FileStream("HouseInfo.f_IDRES.config", FileMode.Create))
                {
                    formatter2.Serialize(fs, settings_IDRES);
                }
                Console.WriteLine("Файлыс насроек созданы.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }          
        }

        static void OpenSettings(out ExcelHandler.Settings.Settings_109 settings_109, out ExcelHandler.Settings.Settings_IDRES settings_IDRES)
        {
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(HouseInfo.ExcelHandler.Settings.Settings_109));
                using (FileStream fs = new FileStream("HouseInfo.f_109.config", FileMode.Open))
                {
                    settings_109 = (HouseInfo.ExcelHandler.Settings.Settings_109)formatter.Deserialize(fs);
                }

                Console.WriteLine("Настройки из 'HouseInfo.f_109.config' загружены.");
            }
            catch (Exception ex)
            {
                settings_109 = new ExcelHandler.Settings.Settings_109(true);
                Console.WriteLine(ex.Message);
                Console.WriteLine("Использованы базовые натройки.");
            }
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(HouseInfo.ExcelHandler.Settings.Settings_IDRES));
                using (FileStream fs = new FileStream("HouseInfo.f_IDRES.config", FileMode.Open))
                {
                    settings_IDRES = (HouseInfo.ExcelHandler.Settings.Settings_IDRES)formatter.Deserialize(fs);
                }

                Console.WriteLine("Настройки из 'HouseInfo.f_IDRES.config' загружены.");
            }
            catch (Exception ex)
            {
                settings_IDRES = new ExcelHandler.Settings.Settings_IDRES(true);
                Console.WriteLine(ex.Message);
                Console.WriteLine("Использованы базовые натройки.");
            }          
        }

        static void ShowMassage(string msg)
        {
            Console.Clear();
            Console.WriteLine(msg);
        }

   
    }
}
