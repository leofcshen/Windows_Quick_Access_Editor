using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Windows_Quick_Access_Editor
{
    class CMyClass
    {
        private string _configPath = "\\config.ini";
        public string configPath { get; } // config.ini path
        public static Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
        public static Object shell = Activator.CreateInstance(shellAppType);
        public Shell32.Folder2 f2 = (Shell32.Folder2)shellAppType.InvokeMember
                ("NameSpace", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}" });
        public string[] defaultQuickAccess = new string[4] { "桌面", "下載", "文件", "圖片" };
        public class CDicValue
        {
            public string Name { get; set; }
            public string Path { get; set; }
        }
        /// <summary>
        /// 目前的快速存取清單
        /// </summary>
        public Dictionary<int, CDicValue> dCurrentList { get; set; }
        public CMyClass()
        {
            mAddQuickAccessToDic();
            configPath = Directory.GetCurrentDirectory() + _configPath;
        }
        
        /// <summary>
        /// 確認 config.ini 是否存在
        /// </summary>
        public bool mCheckConfigFile()
        {            
            if (File.Exists(configPath))
                return true;
            else
                return false;
        }

        /// <summary>
        /// 新增 config.ini
        /// </summary>
        /// <returns>string 執行結果</returns>
        public string mAddConfigFile() 
        {
            try
            {
                using (FileStream fs = new FileStream(configPath, FileMode.Create, FileAccess.Write))
                {
                    using (StreamWriter sr = new StreamWriter(fs))
                    {
                        foreach (var item in dCurrentList)
                            sr.WriteLine(item.Value.Path);
                    }
                }
                return "config.ini 檔已建立，請編輯後使用。";
            }
            catch (Exception e) { return $"config.ini 建立失敗，錯誤碼 {e.Message}"; }            
        }

        /// <summary>
        /// 刪除自訂快速存取
        /// </summary>
        /// <returns>string 執行結果</returns>
        public string mDelCustomer()
        {            
            // 路徑在 "%AppData%\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms"
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms";
            try
            {
                File.Delete(filePath);
                return "刪除自訂快速存取檔案完成。";
            }
            catch (Exception e) { return "刪除自訂快速存取失敗，錯誤碼：" + e.Message; }
        }

        /// <summary>
        /// 刪除單筆快速存取
        /// </summary>
        /// <returns>string 執行結果</returns>
        public string mDelSingle(string delFolder)
        {
            foreach (Shell32.FolderItem fi in f2.Items())
            {
                if (fi.Name == delFolder)
                {
                    ((Shell32.FolderItem)fi).InvokeVerb("unpinfromhome");
                    break;
                }
            }
            return "刪除單筆完成：" + delFolder;
        }

        /// <summary>
        /// 把快速存取清單加到 dCurrentList
        /// </summary>
        public void mAddQuickAccessToDic()
        {
            f2 = (Shell32.Folder2)shellAppType.InvokeMember
                ("NameSpace", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}" });
            dCurrentList = new Dictionary<int, CDicValue>();
            int count = 1;
            foreach (Shell32.FolderItem fi in f2.Items())
            {
                CDicValue cDicValue = new CDicValue { Name = fi.Name, Path = fi.Path};
                dCurrentList.Add(count, cDicValue);
                count++;
            }
        }

        /// <summary>
        /// 新增快速存取
        /// </summary>   
        /// <returns>string 執行結果</returns>
        public string mAddQuickAccess(List<string> lisFolder)
        {
            string str = string.Empty;

            foreach (var item in lisFolder)
            {
                f2 = (Shell32.Folder2)shellAppType.InvokeMember
                    ("NameSpace", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { item });
                try
                {
                    f2.Self.InvokeVerb("pintohome");
                    str += $"新增 {item} 成功。\n";
                }
                catch (Exception e) { str += $"新增 {item} 失敗，錯誤碼：{e.Message}。\n"; }
            }
            return str;
        }

        /// <summary>
        /// 讀取 config.ini 傳回 configList
        /// </summary>
        /// <returns>bool 成敗</returns>
        public bool mReadConfig(out List<string> configList)
        {
            configList = new List<string>();
            try
            {
                using (StreamReader sr = new StreamReader(configPath, System.Text.Encoding.UTF8))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                        configList.Add(line);
                }
                return true;
            }
            catch (Exception ex) { return false; }
        }
        /// <summary>
        /// 匯出快速存取到 config.ini
        /// </summary>
        /// <returns>string 執行結果</returns>
        public string mExportConfig()
        {
            try
            {
                using (FileStream fs = new FileStream(configPath, FileMode.Create, FileAccess.Write))
                {
                    using (StreamWriter sr = new StreamWriter(fs))
                    {
                        foreach (var item in dCurrentList)
                            sr.WriteLine(item.Value.Path);
                    }
                }
                return "config.ini 已匯出。";
            }
            catch (Exception e) { return $"config.ini 匯出失敗，錯誤碼：{e.Message}。"; }
        }

        /// <summary>
        /// 刪除所有快速存取
        /// </summary>
        public string mDelAll()
        {
            foreach (Shell32.FolderItem fi in f2.Items())
            {
                ((Shell32.FolderItem)fi).InvokeVerb("unpinfromhome");
            }
            return "刪除所有快速存取已完成。";
        }

        /// <summary>
        /// 新增預設快速存取
        /// </summary>
        public string mAddDefault()
        {
            List<string> lisFolder = new List<string>() {
                @"C:\Users\rovin\Desktop",
                @"C:\Users\rovin\Downloads",
                @"C:\Users\rovin\Documents",
                @"C:\Users\rovin\Pictures",
            };
            string str = string.Empty;

            foreach (var item in lisFolder)
            {
                f2 = (Shell32.Folder2)shellAppType.InvokeMember
                    ("NameSpace", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { item });
                try
                {
                    f2.Self.InvokeVerb("pintohome");
                    str += $"新增 {item} 成功。\n";
                }
                catch (Exception e) { str += $"新增 {item} 失敗，錯誤碼：{e.Message}。\n"; }
            }
            return str;
        }
    }
}
