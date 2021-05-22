using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
// 加入參考：Microsoft Shell Controls and Automation

namespace Windows_Quick_Access_Editor
{
    class Program
    {
        static void Main(string[] args)
        {
            bool run = true;
            do
            {
                Console.WriteLine("1.刪除所有快速存取");
                Console.WriteLine("2.批次新增快速存取");
                Console.WriteLine("3.單筆新增快速存取");
                Console.WriteLine("4.離開");

                bool success = Int32.TryParse(Console.ReadLine(), out int number);

                if (success)
                {
                    switch (number)
                    {
                        case 1: // 刪除所有快速存取
                            // 路徑在 "%AppData%\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms"
                            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms";
                            if (File.Exists(filePath)) // 檔案存在的話刪除
                            {
                                try
                                {
                                    File.Delete(filePath);
                                    Console.WriteLine("刪除完成");
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("刪除失敗，錯誤碼：" + e.Message);
                                }
                            }
                            else
                                Console.WriteLine("檔案不存在");
                            break;

                        case 2: // 批次新增快速存取
                            string configPath = Directory.GetCurrentDirectory() + "\\config.ini";
                            List<string> list = new List<string>();

                            try
                            {
                                if (File.Exists(configPath)) // config.ini 存在，把路徑存進 list。
                                {
                                    using (StreamReader sr = new StreamReader(configPath, System.Text.Encoding.Default))
                                    {
                                        string line;
                                        while ((line = sr.ReadLine()) != null)
                                            list.Add(line);
                                    }
                                    foreach (var item in list)
                                        addQuickAccess(item);
                                }
                                else // config.ini 不存在，建立檔案，寫入 default 路徑。
                                {
                                    using (FileStream fs = new FileStream(configPath, FileMode.Create, FileAccess.Write))
                                    {
                                        using (StreamWriter sr = new StreamWriter(fs))
                                        {
                                            sr.WriteLine("C:\\ProgramData");
                                            sr.WriteLine("C:\\Windows");
                                        }
                                    }
                                    Console.WriteLine("config.ini 檔已建立，請編輯後使用");
                                }
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("批次新增失敗，錯誤碼：" + e.Message);
                            }
                            break;

                        case 3: // 單筆新增快速存取
                            Console.WriteLine("請輸入要加入快速存取的資料夾路徑：(例如 'D:\\PC')");
                            string folderPath = Console.ReadLine();
                            addQuickAccess(folderPath);
                            break;

                        case 4: // 離開
                            run = false;
                            break;
                    }
                }
                separator();
            } while (run);

            void addQuickAccess(string folderPath)
            {
                try
                {
                    Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
                    Object shell = Activator.CreateInstance(shellAppType);
                    Shell32.Folder2 f = (Shell32.Folder2)shellAppType.InvokeMember
                        ("NameSpace", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { folderPath });
                    f.Self.InvokeVerb("pintohome");
                    Console.WriteLine("新增成功：" + folderPath);
                }
                catch (Exception e)
                {
                    Console.WriteLine("新增失敗：" + folderPath + " 錯誤碼：" + e.Message);
                }
            }
            void separator() // 分隔線
            {
                string str = new string('=', 20);
                Console.WriteLine(str);
            }

            Console.WriteLine("運行結束，按任意鍵繼續");
            Console.Read();
        }
    }
}