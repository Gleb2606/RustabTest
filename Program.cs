using System.IO;
using System.Security.Cryptography;

namespace RustabTest 
{
    class Program 
    {
        public static void Main(string[] args) 
        {
            // Шаблон файла динамики
            string shablRst = @"C:\Users\geu2\Documents\RastrWin3\SHABLON\динамика.rst";

            // Шаблон файла сценария
            string shablScn = @"C:\Users\geu2\Documents\RastrWin3\SHABLON\сценарий.scn";

            // Шаблон файла автоматики
            string shablDfw = @"C:\Users\geu2\Documents\RastrWin3\SHABLON\автоматика.dfw";

            // Шаблон файла контролируемых величин
            string shablKpr = @"C:\Users\geu2\Documents\RastrWin3\SHABLON\контр-е величины.kpr";

            // Путь сохранения режимов
            string saveRstPath = "C:\\Users\\geu2\\Desktop\\rst";

            // Путь сохранения сценариев
            string saveScnPath = "C:\\Users\\geu2\\Desktop\\scn";

            // Путь сохранения результатов расчета
            string saveResultPath = @"C:\test";

            // Исходный файл
            string file = @"C:\Users\geu2\Desktop\ВКР  temp (1)\Модель БоГЭС\БоГЭС.rst";

            // Загружаем файл режима
            RastrSupplier.LoadFile(file, shablRst);

            // Пример для БоАЗ
            int lowerBound = 252;
            int upperBound = 307;
            int increment = 5;
            int BoAZNode = 60534001;

            // В диапазоне изменения влияющих факторов
            while (lowerBound < upperBound)
            {
                // Изменить значение
                RastrSupplier.SetValue("node", "ny", BoAZNode, "pn", lowerBound);

                // Проверка сходимости режима
                bool isRegimeOK = RastrSupplier.IsRegimeOK();

                // Если режим сошёлся
                if (isRegimeOK)
                {
                    Console.WriteLine("Режим сошёлся, сохраняем файл");

                    // Сохраняем файл
                    RastrSupplier.SaveFile($"{saveRstPath}\\БоАЗ_{lowerBound}.rst",
                        shablRst);
                }
                else 
                {
                    Console.WriteLine("Режим не сходится");
                }

                // Приращение
                lowerBound += increment;
            }

            // Пример для возмущения I группы
            double shunt = 9.297;
            double shuntRecloser = 11.597;
            double protectionDelay = 0.08;
            double backupDelay = 0.1;
            double recloserDelay = 5.195;
            int faultNode = 60401001;
            int newFaultNode = 60533056;
            string line1 = "60533027,60533056,0";
            string line2 = "60533003,60533056,0";
            string line3 = "60401001,60401005,0";
            string line4 = "60401001,60401142,0";

            //// Формирование сценариев
            //for (int i = 1; i < 4; i++) 
            //{
            //    // Загружаем файл сценария
            //    RastrSupplier.LoadFile(shablScn, shablScn);

            //    switch (i) 
            //    {
            //        // Исходный сценарий
            //        case 1:
            //        {
            //             RastrSupplier.MakeScnI(shunt, shuntRecloser,
            //                 protectionDelay, recloserDelay, faultNode,
            //                 newFaultNode, line1, line2, line3, line4);
            //             break;
            //        }

            //        // Переходное сопротивление
            //        case 2: 
            //        {
            //            RastrSupplier.MakeScnI(shunt + 5, shuntRecloser + 5,
            //                protectionDelay, recloserDelay, faultNode,
            //                newFaultNode, line1, line2, line3, line4);
            //            break;
            //        }

            //        // Отключение резервной защитой
            //        case 3:
            //        {
            //            RastrSupplier.MakeScnI(shunt, shuntRecloser,
            //                backupDelay, recloserDelay, faultNode,
            //                newFaultNode, line1, line2, line3, line4);
            //            break;
            //        }
            //    }
            //    Console.WriteLine("Сценарий сформирован, сохраняем файл");

            //    RastrSupplier.SaveFile($"{saveScnPath}\\сценарий_{i}.scn",
            //        shablScn);
            //}

            // Пример для возмущения III группы
            double shuntIII = 47.406;
            double CBFPDelay = 0.255;
            int faultNodeIII = 60533060;
            string line1III = "60533001,60533060,0";
            string line2III = "60512158,60512163,0";
            string line3III = "60512163,60512164,0";
            string line4III = "60533027,60533059,0";
            string line5III = "60533027,60533056,0";

            // Формирование сценариев
            for (int i = 1; i < 4; i++)
            {
                // Загружаем файл сценария
                RastrSupplier.LoadFile(shablScn, shablScn);

                switch (i)
                {
                    // Исходный сценарий
                    case 1:
                        {
                            RastrSupplier.MakeScnIII(shuntIII, protectionDelay,
                                CBFPDelay, faultNodeIII, line1III, line2III,
                                line3III, line4III, line5III);
                            break;
                        }

                    // Переходное сопротивление
                    case 2:
                        {
                            RastrSupplier.MakeScnIII(shuntIII + 5,
                                protectionDelay, CBFPDelay, faultNodeIII,
                                line1III, line2III, line3III, line4III,
                                line5III);
                            break;
                        }

                    // Отключение резервной защитой
                    case 3:
                        {
                            RastrSupplier.MakeScnIII(shuntIII, backupDelay,
                                CBFPDelay, faultNodeIII, line1III, line2III,
                                line3III, line4III, line5III);
                            break;
                        }
                }
                Console.WriteLine("Сценарий сформирован, сохраняем файл");

                RastrSupplier.SaveFile($"{saveScnPath}\\сценарий_УРОВ_{i}.scn",
                    shablScn);
            }

            // Пример КВ Богучанская ГЭС - Красноярская ГЭС
            string name = "Богучанская ГЭС";
            int index = 60533014;
            string set = "Num=60533014";
            string nameSupport = "Красноярская ГЭС";
            string setSupport = "Num=60522003";
            
            // Пакетный расчет динамики
            try
            {
                // Получение файлов динамики и сценариев
                string[] rstFiles = Directory.GetFiles(saveRstPath);
                string[] scnFiles = Directory.GetFiles(saveScnPath);
                int resIndex = 1;
                int snapIndex = 0;

                // Список потоков
                List<Task> tasks = new List<Task>();

                foreach (string rst in rstFiles)
                {
                    foreach (string scn in scnFiles)
                    {
                        
                        // Подготовка файлов динамики
                        RastrSupplier.FilePrepair(rst, scn, shablDfw, shablScn,
                        shablKpr, name, set, nameSupport, setSupport);

                        // Многопоточное вычисление
                        tasks.Add(Task.Run(() =>
                        {
                            // Расчет динамики
                            RastrSupplier.CalculateDynamic(5, saveResultPath);

                            // Сохранение результатов
                            RastrSupplier.GetTransient("Generator", "Delta", index,
                                snapIndex, "БоГЭС", resIndex);
       
                            resIndex++;
                            snapIndex++;
                        }));
                    }
                }

                // Ожидание завершения всех задач
                Task.WaitAll(tasks.ToArray());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
