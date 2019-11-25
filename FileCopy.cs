using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.IO;

namespace FileCopy
{
    public partial class Form1 : Form
    {
        SynchronizationContext _syncContext;

        bool DeleteSwitch = false; //檢查當天是否刪除過檔案了        
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _syncContext = SynchronizationContext.Current;

            Thread Thread_A = new Thread(new ThreadStart(WorkA));
            Thread_A.IsBackground = true;
            Thread_A.Start();

            Thread Thread_B = new Thread(new ThreadStart(WorkB));
            Thread_B.IsBackground = true;
            Thread_B.Start();
        }

        public void WorkA()
        {
            //2小時複製檔案一次
            while (true)
            {
                SupPuzzle();
                Thread.Sleep(7200000);
            }
        }

        public void WorkB()
        {            
            //檢查是否要刪除檔案
            while (true)
            {
                DateTime NowDate = DateTime.Now; //現在時間

                if (NowDate.Hour >= 23 || NowDate.Hour <= 2)
                {
                    DeleteSwitch = true;
                }

                if (NowDate.Hour == 12 || NowDate.Hour == 18)
                {
                    PuzzleTouch();
                }

                Thread.Sleep(3600000);
            }
        }

        //超級吞T
        private void SupPuzzle()
        {
            DateTime NowDate = DateTime.Now; //現在時間
            string SupPuzzleSource = @"D:\超級吞T"; //原檔位置
            string SupPuzzleDirectory = @"D:\備份檔案\SupPuzzle\" + NowDate.ToString("yyyy年MM月dd日") + "\\" + NowDate.ToString("MM月dd日HH時"); //傳輸目標
            string DeleteSupPuzzle = @"D:\備份檔案\SupPuzzle\" + NowDate.AddDays(-3).ToString("yyyy年MM月dd日");//刪除檔案夾

            #region 複製檔案

            try
            {
                DirectoryInfo dir = new DirectoryInfo(SupPuzzleSource);

                if (NowDate.Hour <= 22 && NowDate.Hour >= 8) //早上8點到晚上22點備份檔案
                {
                    if (dir.Exists)
                    {

                        DirectoryCopy(SupPuzzleSource, SupPuzzleDirectory, true);
                        _syncContext.Post(Result, SupPuzzleDirectory + " 複製成功");

                    }
                    else
                    {
                        _syncContext.Post(Result, SupPuzzleSource + " 要複製檔案夾不存在");
                    }
                }
            }
            catch
            {
                _syncContext.Post(Result, SupPuzzleSource + " 複製失敗");
            }

            #endregion

            #region 刪除檔案

            if (DeleteSwitch && NowDate.Hour >= 6 && NowDate.Hour <= 9) //早上6點到9點刪除前三天檔案
            {
                try
                {
                    DirectoryInfo dir = new DirectoryInfo(DeleteSupPuzzle);

                    if (dir.Exists)
                    {
                        DeleteFolder(DeleteSupPuzzle);
                        DeleteSwitch = false;
                        _syncContext.Post(Result, DeleteSupPuzzle + " 刪除成功");
                    }
                }
                catch
                {
                    _syncContext.Post(Result, DeleteSupPuzzle + " 刪除失敗");
                }
            }

            #endregion     
        }

        //吞T
        private void PuzzleTouch()
        {
            DateTime NowDate = DateTime.Now; //現在時間
            string PuzzleSource = @"\\pc7361\更新丟檔區"; //原檔位置
            string PuzzleDirectory = @"D:\備份檔案\Puzzle\" + NowDate.ToString("yyyy年MM月dd日") + "\\" + NowDate.ToString("MM月dd日HH時"); //傳輸目標
            string DeletePuzzle = @"D:\備份檔案\Puzzle\" + NowDate.AddDays(-3).ToString("yyyy年MM月dd日");//刪除檔案夾  

            #region 複製檔案

            try
            {
                DirectoryInfo dir = new DirectoryInfo(PuzzleSource);

                if (dir.Exists)
                {
                    DirectoryCopy(PuzzleSource, PuzzleDirectory, true);
                    _syncContext.Post(Result, PuzzleDirectory + " 複製成功");
                }
                else
                {
                    _syncContext.Post(Result, PuzzleSource + " 要複製檔案夾不存在");
                }
            }
            catch
            {
                _syncContext.Post(Result, PuzzleSource + " 複製失敗");
            }

            #endregion

            #region 刪除檔案

            if (DeleteSwitch && NowDate.Hour >= 6 && NowDate.Hour <= 9) //早上6點到9點刪除前三天檔案
            {
                try
                {
                    DirectoryInfo dir = new DirectoryInfo(DeletePuzzle);

                    if (dir.Exists)
                    {
                        DeleteFolder(DeletePuzzle);
                        DeleteSwitch = false;
                        _syncContext.Post(Result, DeletePuzzle + " 刪除成功");
                    }
                }
                catch
                {
                    _syncContext.Post(Result, DeletePuzzle + " 刪除失敗");
                }
            }

            #endregion
        }

        //資訊顯示
        int Times = 0; //資訊顯示次數
        private void Result(object result) 
        {
            DateTime NowDate = DateTime.Now; //現在時間

            try
            {
                if (Times <= 100)
                {
                    richTextBox_Display.Text += result.ToString() + "---" + NowDate.ToString("yyyy-MM-dd HH:mm:ss") + "\n";
                    Times++;
                }
                else
                {
                    richTextBox_Display.Text = "";
                    Times = 0;
                    richTextBox_Display.Text += result.ToString() + "---" + NowDate.ToString("yyyy-MM-dd HH:mm:ss") + "\n";
                    Times++;
                }
            }
            catch
            {
                richTextBox_Display.Text += "顯示記錄失敗" + "\n";
                Times++;
            }

        }

        //複製檔案詳細
        int CopyFileTimes = 0; //資訊顯示次數
        private void CopyFileResult(object result)
        {
            DateTime NowDate = DateTime.Now; //現在時間

            try
            {
                if (CopyFileTimes <= 20)
                {
                    richTextBox_CopyFileDetail.Text += result.ToString() + "---" + NowDate.ToString("yyyy-MM-dd HH:mm:ss") + "\n";
                    CopyFileTimes++;
                }
                else
                {
                    richTextBox_CopyFileDetail.Text = "";
                    CopyFileTimes = 0;
                    richTextBox_CopyFileDetail.Text += result.ToString() + "---" + NowDate.ToString("yyyy-MM-dd HH:mm:ss") + "\n";
                    CopyFileTimes++;
                }
            }
            catch
            {
                richTextBox_CopyFileDetail.Text += "顯示記錄失敗" + "\n";
                CopyFileTimes++;
            }

        }     

       
        //Function區
        /*----------------------------------------------------------------------------------------------------------------------------------------------------------------*/
        
        private void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs) //複製檔案
        {
            try
            {
                // Get the subdirectories for the specified directory.
                DirectoryInfo dir = new DirectoryInfo(sourceDirName);
                DirectoryInfo[] dirs = dir.GetDirectories();
                                
                // If the destination directory doesn't exist, create it.
                if (!Directory.Exists(destDirName))
                {
                    Directory.CreateDirectory(destDirName);
                }

                // Get the files in the directory and copy them to the new location.
                FileInfo[] files = dir.GetFiles();
                foreach (FileInfo file in files)
                {
                    string temppath = Path.Combine(destDirName, file.Name);
                    file.CopyTo(temppath, true);
                    _syncContext.Post(CopyFileResult, temppath + "--複製成功");
                }

                // If copying subdirectories, copy them and their contents to new location.
                if (copySubDirs)
                {
                    foreach (DirectoryInfo subdir in dirs)
                    {
                        string temppath = Path.Combine(destDirName, subdir.Name);
                        DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                        _syncContext.Post(CopyFileResult, temppath + "--複製成功");
                    }
                }
            }
            catch (IOException e)
            {
                _syncContext.Post(Result, e.Message);
            }
        }

        private void DeleteFolder(string DeleteFolder) //刪除檔案夾
        {
            DirectoryInfo di = new DirectoryInfo(DeleteFolder);

            try
            {
                di.Delete(true);
            }

            catch (IOException e)
            {
                _syncContext.Post(Result, e.Message);
            }
        }

        private void CreatFolder() //新增資料夾
        {
            // Specify a name for your top-level folder.
            string folderName = @"C:\Users\howerlai1990\Desktop\Test";

            // To create a string that specifies the path to a subfolder under your 
            // top-level folder, add a name for the subfolder to folderName.
            string pathString = System.IO.Path.Combine(folderName, "SubFolder");

            System.IO.Directory.CreateDirectory(pathString);

        }        

        private void FileMove() //移動檔案
        {
            //string fileName = "";
            //string destFile = "";
            //string sourcePath = @"C:\Users\howerlai1990\Desktop\fileok";
            //string targetPath = @"C:\Users\howerlai1990\Desktop\Test\SubFolder";

            //if (System.IO.Directory.Exists(sourcePath))
            //{
            //    string[] files = System.IO.Directory.GetFiles(sourcePath);

            //    // Copy the files and overwrite destination files if they already exist.
            //    foreach (string s in files)
            //    {
            //        // Use static Path methods to extract only the file name from the path.
            //        fileName = System.IO.Path.GetFileName(s);
            //        destFile = System.IO.Path.Combine(targetPath, fileName);
            //        System.IO.File.Copy(s, destFile, true);
            //    }
            //}
            //else
            //{
            //    Console.WriteLine("Source path does not exist!");
            //}          
        }
    }
}
