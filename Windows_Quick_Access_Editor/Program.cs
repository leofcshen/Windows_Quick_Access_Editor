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
            string configPath = Directory.GetCurrentDirectory() + "\\config.ini";
            
            if (!File.Exists(configPath)) // config.ini 如果不存在，建立檔案，寫入 default 路徑。
            {
                try
                {
                    using (FileStream fs = new FileStream(configPath, FileMode.Create, FileAccess.Write))
                    {
                        using (StreamWriter sr = new StreamWriter(fs))
                        {
                            sr.WriteLine("C:\\ProgramData");
                            sr.WriteLine("C:\\Windows");
                        }
                    }
                    Console.WriteLine("config.ini 檔已建立，請編輯後使用。");
                }
                catch (Exception e) { Console.WriteLine("config.ini 新增失敗，錯誤碼：" + e.Message); }
            }
            else
                Console.WriteLine("config.ini 檔已存在，請編輯後使用。");
            separator();

            bool run = true; // 迴圈 flag
            bool success; // 測試轉換選項
            int number; // 選項
            do
            {
                Console.WriteLine("1.刪除所有快速存取");
                Console.WriteLine("2.刪除單筆快速存取");
                Console.WriteLine("3.新增批次快速存取");
                Console.WriteLine("4.新增單筆快速存取");                
                Console.WriteLine();
                Console.Write("請選擇功能，輸入 0 離開：");

                success = Int32.TryParse(Console.ReadLine(), out number); // 轉換選項
                Console.WriteLine();

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
                                    Console.WriteLine("刪除全部完成");
                                }
                                catch (Exception e) { Console.WriteLine("刪除全部失敗，錯誤碼：" + e.Message); }
                            }
                            else
                                Console.WriteLine("檔案不存在");
                            break;
                        case 2: // 單筆移除快速存取
                            delQuickAccess();                            
                            break;

                        case 3: // 批次新增快速存取                            
                            List<string> addList = new List<string>(); // 用來存放 config.ini 的路徑

                            try
                            {
                                if (File.Exists(configPath)) // config.ini 存在，把路徑存進 list。
                                {                                    
                                    using (StreamReader sr = new StreamReader(configPath, System.Text.Encoding.Default))
                                    {
                                        string line;
                                        while ((line = sr.ReadLine()) != null)
                                        {
                                            addList.Add(line);
                                            Console.WriteLine(line);
                                        }
                                    }
                                    Console.WriteLine();
                                    Console.Write("確認新增上列快速存取資料夾？(確定：1，取消：0)：");
                                    success = Int32.TryParse(Console.ReadLine(), out number); // 轉換選項
                                    if (success && number == 1)
                                    {
                                        foreach (var item in addList)
                                            addQuickAccess(item);
                                        Console.WriteLine("新增批次成功");
                                    }
                                }
                                else
                                    Console.WriteLine("config.ini 不存在。");
                            }
                            catch (Exception e) { Console.WriteLine("新增批次失敗，錯誤碼：" + e.Message); }
                            break;

                        case 4: // 單筆新增快速存取
                            Console.WriteLine("請輸入要加入快速存取的資料夾路徑：(例如 'D:\\PC')");
                            string addFolderPath = Console.ReadLine();
                            Console.WriteLine();
                            addQuickAccess(addFolderPath);
                            break;

                        case 0: // 離開
                            run = false;
                            break;
                    }
                }
                separator();
            } while (run);

            Console.WriteLine("運行結束，按任意鍵繼續");
            Console.Read();
        }

        static void separator() // 分隔線
        {
            string str = new string('=', 50);
            Console.WriteLine(str);
        }

        static void addQuickAccess(string addFolderPath) // 新增快速存取
        {
            try
            {
                Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
                Object shell = Activator.CreateInstance(shellAppType);
                Shell32.Folder2 f2 = (Shell32.Folder2)shellAppType.InvokeMember
                    ("NameSpace", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { addFolderPath });
                f2.Self.InvokeVerb("pintohome");
                Console.WriteLine("新增單筆成功：" + addFolderPath);
            }
            catch (Exception e) { Console.WriteLine("新增單筆失敗：" + addFolderPath + " 錯誤碼：" + e.Message); }
        }

        static void delQuickAccess() // 移除快速存取
        {
            string delFolderName = String.Empty;

            try
            {
                bool run = true;

                do
                {
                    Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
                    Object shell = Activator.CreateInstance(shellAppType);
                    Shell32.Folder2 f2 = (Shell32.Folder2)shellAppType.InvokeMember("NameSpace", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}" });
                    int count = 1;
                    Dictionary<int, string> dCurrentList = new Dictionary<int, string>(); // 存放目前的快速存存清單

                    foreach (Shell32.FolderItem fi in f2.Items()) // 列出清單、加入 Dictionary
                    {
                        Console.WriteLine(count + "：" + fi.Name);
                        dCurrentList.Add(count, fi.Name);
                        count++;
                    }
                    Console.WriteLine();
                    Console.Write("請選擇要移除的項目，輸入 0 離開：");

                    bool success = Int32.TryParse(Console.ReadLine(), out int number); // 轉換選項
                    Console.WriteLine();
                    if (number == 0) // 離開
                    {
                        run = false;
                        return;
                    }
                    if (dCurrentList.ContainsKey(number)) // 用 key 尋找 Dictionary，找到的話 delFolderName 賦值。
                        delFolderName = dCurrentList[number];
                    else
                    {
                        Console.WriteLine("找不到項目：" + number);
                    }

                    if (success)
                    {
                        foreach (Shell32.FolderItem fi in f2.Items())
                        {
                            if (fi.Name == delFolderName)
                            {
                                ((Shell32.FolderItem)fi).InvokeVerb("unpinfromhome");
                                Console.WriteLine("刪除單筆完成：" + delFolderName);
                            }
                        }
                    }
                    separator();
                } while (run);
            }
            catch (Exception e) { Console.WriteLine("刪除單筆失敗：" + delFolderName + " 錯誤碼：" + e.Message); }
        }
    }
}