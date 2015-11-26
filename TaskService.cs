using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32.TaskScheduler;
namespace TeachTaskService
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Begin
            TaskService ts = CreateTaskService();
            IEnumerable<Task> _result = GetAllTask(ts);
            _result = filterByFolder(_result);
            ShowConnectionState(ts);
            #endregion

            #region Teach
            //A.找出TaskServer指定資料夾中JOB名稱            
            ShowItem(_result);
              
            //B.找出TaskServer中JOB的"下次預計執行時間"
            ShowNextRunTime(_result);
            #endregion

            #region End
            CloseTaskService(ts);
            ShowConnectionState(ts);
            Console.ReadLine();
            #endregion
        }
        #region TeachMethod
        /// <summary>
        /// 展示JOB的名稱、狀態、下次執行時間
        /// </summary>        
        private static void ShowItem(IEnumerable<Task> _result)
        {           
            Console.WriteLine("==================JOB資訊展示開始==================");            
            foreach (var v in _result)
            {
                Console.WriteLine(String.Format("JobName:{0},Status:{1},NextRunTime:{2}", v.Name, v.State, v.NextRunTime));
            }
            Console.WriteLine("==================JOB資訊展示結束==================");            
        }

        /// <summary>
        /// GetRunTimes:取得時間區內所有下次執行時間
        /// </summary>
        private static void ShowNextRunTime(IEnumerable<Task> _result)
        {       
            int ShowCount = 0;
            Console.WriteLine("==================JOB時間區間撈取開始==================");
            foreach (var v in _result)
            {
                foreach (var p in v.GetRunTimes(Utility.sTime.ToUniversalTime(), Utility.eTime.ToUniversalTime()))
                {                    
                    Console.WriteLine(String.Format("JobName:{0},NextRunTime:{1}",v.Name, p));                    
                    ShowCount++;
                    if (ShowCount > 2) //每個JOB最多秀3筆
                    {
                        ShowCount = 0;
                        break;
                    }
                }
            }
            Console.WriteLine("==================JOB時間時間區間撈取結束==================");
        }

        #endregion

        #region Connection
        /// <summary>
        /// 建立連線
        /// TaskService(IP,Account,DOMAIN,Password)
        /// </summary>        
        private static TaskService CreateTaskService()
        {            
            return new TaskService(TaskServiceInfo.ServerName, TaskServiceInfo.Account, TaskServiceInfo.DOMAIN, TaskServiceInfo.Password);            
        }

        /// <summary>
        /// 關閉連線
        /// </summary>        
        private static void CloseTaskService(TaskService ts)
        {
            ts.Dispose();         
        }

        /// <summary>
        /// 判斷連線狀態
        /// ts.Connected
        /// </summary>        
        private static void ShowConnectionState(TaskService ts)
        {
            String _state = ts.Connected ? "TaskServer連線狀態:已連線" : "TaskServer連線狀態:未連線";
            Console.WriteLine(_state);
        }
        #endregion

        #region Data
        /// <summary>
        /// 取得所有TaskServer掛載的JOB
        /// </summary>
        private static IEnumerable<Microsoft.Win32.TaskScheduler.Task> GetAllTask(TaskService ts)
        {
            return ts.AllTasks;
        }
        #endregion       

        #region Transform
        /// <summary>
        /// 取出指定的Folder資料夾下的JOB
        /// Task.Folder.Name找出該Item所屬的Folder名稱
        /// </summary>
        private static IEnumerable<Task> filterByFolder(IEnumerable<Task> _result)
        {
            return _result.Where(p => p.Folder.Name == Utility.Folder);
        }
        #endregion
    }
}
