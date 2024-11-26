using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DownloadWatcher
{
    public partial class Form1 : Form
    {

        public const string AppTitle = "DownloadWatcher";

        static string machineName;
        static string userName;
        static string logFileName;
        static string logPath;
        static string logFilePath;

        static FileSystemWatcher watcher;

        static NotifyIcon icon; // task tray icon

        static string prevDownloadFile = "";

        public Form1()
        {
            InitializeComponent();

        }

        private void setComponents()
        {
            icon = new NotifyIcon();
            icon.Icon = new Icon(Path.GetDirectoryName(Application.ExecutablePath) + @"\app.ico");
            icon.Visible = true;
            icon.Text = "ダウンロード監視";


            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem menuItem = new ToolStripMenuItem();

            menu.Items.AddRange(new ToolStripMenuItem[]{
                new ToolStripMenuItem("終了", null, (s,e)=>{Close_Click();})
            });

            icon.ContextMenuStrip = menu;

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            setComponents();

            SetCurrentVersionRun();


            //イベントをイベントハンドラに関連付ける
            //フォームコンストラクタなどの適当な位置に記述してもよい
            SystemEvents.SessionEnding += new SessionEndingEventHandler(SystemEvents_SessionEnding);

            // Get Session Info 
            machineName = Environment.MachineName;
            userName = Environment.UserName;

            // log file name (Machine-Date.log)
            DateTime dt = DateTime.Now;

            //logFileName = machineName + "-" + dt.ToString("yyyyMMdd") + ".log";
            logFileName = dt.ToString("yyyyMMdd") + ".log";

            logPath = Environment.GetEnvironmentVariable("USERPROFILE") + "\\Documents\\" + AppTitle;
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            logFilePath = logPath + "\\" + logFileName;


            string downloadFolder = Environment.GetEnvironmentVariable("USERPROFILE") + @"\Downloads";

            Console.WriteLine("watch folder=" + downloadFolder);

            watcher = new FileSystemWatcher();

            watcher.Path = downloadFolder;
            watcher.Filter = "*.*";  // これだとうまく動作する
            watcher.IncludeSubdirectories = true;
            // 監視パラメータの設定
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite | NotifyFilters.Size;

            // イベントハンドラの設定
            //watcher.Created += new FileSystemEventHandler(watcher_Created);
            watcher.Changed += new FileSystemEventHandler(watcher_Renamed);
            watcher.Error += new ErrorEventHandler(watcher_Error);

            //WindowFormなどUI用(コンソールでは不要)
            watcher.SynchronizingObject = this;

            //監視を開始する
            watcher.EnableRaisingEvents = true;

            WriteLog("started");

            ShowNotification("常駐開始しました");

            // 最小化,タスクバー非表示
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
        }


        //バルーン通知
        static void ShowNotification(string Msg)
        {
            //バルーンヒントのタイトル
            icon.BalloonTipTitle = AppTitle;
            //バルーンヒントに表示するメッセージ
            icon.BalloonTipText = Msg;
            //バルーンヒントに表示するアイコン
            icon.BalloonTipIcon = ToolTipIcon.Info;
            //バルーンヒントを表示する
            //表示する時間をミリ秒で指定する
            icon.ShowBalloonTip(10000);

        }



        static void watcher_Renamed(object source, FileSystemEventArgs e)
        {

            // ダウンロードの時って.tmpから正常なファイル拡張子に変更される, 2回発生する
            string fileName;
            fileName = e.Name;

            if (Path.GetExtension(fileName) == ".exe")
            {
                if (prevDownloadFile != fileName)
                {
                    WriteLog("created " + fileName);
                    ShowNotification("exeファイルがダウンロードされました");

                }
            }
            prevDownloadFile = fileName;

        }
        static void watcher_Error(object source, ErrorEventArgs e)
        {
            Console.WriteLine("called watcher Error");
        }

        private static void WriteLog(string logMessage)
        {

            Console.WriteLine(logMessage);
            File.AppendAllText(logFilePath, GetNowTime() + logMessage + "\n");
        }

        // ログファイル用
        private static string GetNowTime()
        {
            DateTime dt = DateTime.Now;
            return dt.ToString("yyyy/MM/dd HH:mm:ss ");
        }

        //ログオフ、シャットダウンをログに出力する
        private void SystemEvents_SessionEnding(object sender, SessionEndingEventArgs e)
        {
            if (e.Reason == SessionEndReasons.Logoff)
            {
                WriteLog("logoff");
            }
            else if (e.Reason == SessionEndReasons.SystemShutdown)
            {
                WriteLog("shutdown");
            }

            //監視を終了
            watcher.EnableRaisingEvents = false;
            watcher.Dispose();
            watcher = null;
        }

        private void Close_Click()
        {
            WriteLog("app exit");

            //イベントを解放する
            //フォームDisposeメソッド内の基本クラスのDisposeメソッド呼び出しの前に
            //記述してもよい
            SystemEvents.SessionEnding -=
                new SessionEndingEventHandler(SystemEvents_SessionEnding);

            Application.Exit();

        }
        public static void SetCurrentVersionRun()
        {
            //Runキーを開く
            Microsoft.Win32.RegistryKey regkey =
                Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Run", true);
            //値の名前に製品名、値のデータに実行ファイルのパスを指定し、書き込む
            regkey.SetValue(Application.ProductName, Application.ExecutablePath);
            //閉じる
            regkey.Close();
        }
    }
}

