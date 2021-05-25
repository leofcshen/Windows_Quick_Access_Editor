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
            string strConfirm = "(確定：1、任意鍵取消)：";
            CMyClass cMC = new CMyClass();

            if (!cMC.mCheckConfigFile())
                Console.WriteLine(cMC.mAddConfigFile());
            else
                Console.WriteLine("config.ini 檔已存在，請編輯後使用。");
            mNewLine();
            mListConfig(cMC);
            mNewLine();
            mListCurrent(cMC);
            mSeparator();
                        
            do
            {
                string defaultItem = string.Empty;
                for (int i = 0; i < cMC.defaultQuickAccess.Length; i++)
                {
                    defaultItem += $"{cMC.defaultQuickAccess[i]}";
                    if (i != cMC.defaultQuickAccess.Length - 1)
                        defaultItem += "、";
                }
                Console.WriteLine($"[預設項目：{defaultItem}]");
                mNewLine();
                Console.WriteLine("1.刪除自訂快速存取(不含預設)");
                Console.WriteLine("2.刪除單筆快速存取");
                Console.WriteLine("3.新增批次快速存取");
                Console.WriteLine("4.新增單筆快速存取");
                Console.WriteLine("5.匯出所有快速存取");
                Console.WriteLine("6.刪除所有快速存取");
                Console.WriteLine("7.新增預設快速存取");
                mNewLine();
                Console.Write("請選擇功能，輸入 0 離開：");

                bool success = Int32.TryParse(Console.ReadLine(), out int option);
                mNewLine();

                if (success)
                {
                    switch (option)
                    {
                        case 1: // 刪除自訂快速存取
                            Console.Write("確認刪除自訂快速存取？" + strConfirm);
                            success = Int32.TryParse(Console.ReadLine(), out option);
                            mNewLine();
                            if (success && option == 1)
                            {
                                Console.WriteLine(cMC.mDelCustomer());
                                mNewLine();
                                mListCurrent(cMC);
                            }
                            else
                                mCancel();
                            break;

                        case 2: // 刪除單筆快速存取
                            do
                            {
                                mListCurrent(cMC);                                
                                mNewLine();

                                Console.Write("請選擇要移除的項目，輸入 0 離開：");
                                success = Int32.TryParse(Console.ReadLine(), out option);
                                mNewLine();

                                if (success)
                                {
                                    if (option == 0) // 離開
                                        break;
                                    if (cMC.dCurrentList.ContainsKey(option)) // 用 key 尋找 Dictionary，找到的話刪除。
                                    {
                                        string delFolder = cMC.dCurrentList[option].Name;
                                        if (option > 0 && option < 5) // 1-4 為系統預設
                                        {
                                            Console.Write("確認移除系統預設項目？" + strConfirm);
                                            success = Int32.TryParse(Console.ReadLine(), out option);
                                            mNewLine();
                                            if (success && option == 1)
                                                Console.WriteLine(cMC.mDelSingle(delFolder));
                                            else 
                                                mCancel();                                            
                                        }
                                    }
                                    else
                                        Console.WriteLine("找不到項目：" + option);
                                }
                                mNewLine();
                            } while (true);
                            break;

                        case 3: // 新增自訂快速存取                           
                            try
                            {
                                if (cMC.mCheckConfigFile())
                                {
                                    if (cMC.mReadConfig(out List<string> addList))
                                    {
                                        Console.WriteLine("[config.ini]");
                                        foreach (var item in addList)
                                            Console.WriteLine(item);
                                    }
                                    else Console.WriteLine("config.ini 讀取失敗。");
                                    mNewLine();

                                    Console.Write("確認新增上列快速存取資料夾？" + strConfirm);
                                    success = Int32.TryParse(Console.ReadLine(), out option);
                                    mNewLine();
                                    if (success && option == 1)
                                        Console.WriteLine(cMC.mAddQuickAccess(addList));
                                    else
                                        mCancel();
                                }
                                else
                                    Console.WriteLine("config.ini 不存在。");
                                mListCurrent(cMC);
                            }
                            catch (Exception e) { Console.WriteLine("新增批次失敗，錯誤碼：" + e.Message); }
                            break;

                        case 4: // 新增單筆快速存取
                            Console.WriteLine("請輸入要加入快速存取的資料夾路徑：(例如 'D:\\PC')");                            
                            List<string> lisFolder = new List<string> { Console.ReadLine() };
                            mNewLine();
                            Console.WriteLine(cMC.mAddQuickAccess(lisFolder));
                            mListCurrent(cMC);
                            break;

                        case 5: // 匯出目前快速存取到 config.ini
                            #region 列出目前的快速存取清單
                            cMC.mAddQuickAccessToDic();
                            foreach (var item in cMC.dCurrentList)
                                Console.WriteLine(item.Key + "." + item.Value.Name);
                            mNewLine();
                            #endregion

                            Console.Write("確認匯出到 config.ini？" + strConfirm);
                            success = Int32.TryParse(Console.ReadLine(), out option);
                            mNewLine();
                            if (success && option == 1)
                            {
                                Console.WriteLine(cMC.mExportConfig());
                                mNewLine();
                                mListConfig(cMC);
                            }
                            else
                                mCancel();
                            break;

                        case 6: // 刪除所有快速存取
                            Console.Write("確認刪除所有快速存取？" + strConfirm);
                            success = Int32.TryParse(Console.ReadLine(), out option);
                            mNewLine();
                            if (success && option == 1)
                            {
                                Console.WriteLine(cMC.mDelAll());
                                mNewLine();
                                mListCurrent(cMC);
                            }
                            else
                                mCancel();
                            break;

                        case 7: // 新增預設快速存取
                            Console.WriteLine(cMC.mAddDefault());
                            Console.WriteLine("新增預設快速存取成功");
                            mNewLine();
                            mListCurrent(cMC);
                            break;

                        case 0: // 離開
                            return;
                    }
                }
                mSeparator();
            } while (true);
        }

        /// <summary>
        /// 列出目前的快速存取清單
        /// </summary>
        static void mListCurrent(CMyClass cMC)
        {
            cMC.mAddQuickAccessToDic();
            Console.WriteLine("[目前的快速存取清單]");
            foreach (var item in cMC.dCurrentList)
                Console.WriteLine(item.Key + "." + item.Value.Name);
        }

        /// <summary>
        /// 列出 config.ini
        /// </summary>
        static void mListConfig(CMyClass cMC)
        {
            Console.WriteLine("[config.ini]");

            cMC.mReadConfig(out List<string> configList);
            foreach (var item in configList)
                Console.WriteLine(item);
        }

        static void mSeparator() // 分隔線
        {
            string str = new string('=', 50);
            Console.WriteLine(str);
        }

        static void mNewLine()
        {
            Console.WriteLine();
        }
        
        static void mCancel()
        {
            Console.WriteLine("已取消");
        }
    }
}