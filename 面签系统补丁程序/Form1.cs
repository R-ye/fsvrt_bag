using IWshRuntimeLibrary;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace 面签系统补丁程序
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string softpath = "";
        /// <summary>
        /// 从资源文件中抽取资源文件
        /// </summary>
        /// <param name="resFileName">资源文件名称（资源文件名称必须包含目录，目录间用“.”隔开,最外层是项目默认命名空间）</param>
        /// <param name="outputFile">输出文件</param>
        public void ExtractResFile(string resFileName, string outputFile)
        {
            BufferedStream inStream = null;
            FileStream outStream = null;
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly(); //读取嵌入式资源
                inStream = new BufferedStream(asm.GetManifestResourceStream(resFileName));
                outStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write);

                byte[] buffer = new byte[1024];
                int length;

                while ((length = inStream.Read(buffer, 0, buffer.Length)) > 0)
                {
                    outStream.Write(buffer, 0, length);
                }
                outStream.Flush();
            }
            finally
            {
                if (outStream != null)
                {
                    outStream.Close();
                }
                if (inStream != null)
                {
                    inStream.Close();
                }
            }
        }

        [DllImport("kernel32.dll")]

        public static extern uint WinExec(string lpCmdLine, uint uCmdShow);

        private void button1_Click(object sender, EventArgs e)
        {
            string desktopPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
          listBox1.Items.Add( "检测到当前系统桌面位置：" + desktopPath);
            string[] xx = Directory.GetDirectories(desktopPath);
            string[] path1 = Directory.GetFiles(desktopPath, "*.lnk");
        
            List<FilePathModel> lstAllFile = new List<FilePathModel>();
            if (path1 != null && path1.Length > 0)
            {
                int t = 0;
                int succ = 0;
                foreach (var path in path1)
                {
                    FilePathModel temp = new FilePathModel();

                    temp.Name = Path.GetFileName(path);
                    temp.Filepath = path;
                    WshShell shell = new WshShell();
                    IWshShortcut lnkPath = (IWshShortcut)shell.CreateShortcut(path);

                    temp.FileTargetPath = lnkPath.TargetPath;
                    //lstAllFile.Add(temp);
                    if (temp.Name.IndexOf("面签录音录像系统") != -1 && lnkPath.Description=="fsvrt")
                    {
                        t = 1;
                        int tag = 0;
                        softpath = temp.FileTargetPath;
                        listBox1.Items.Add( "已找到目标程序位置 ："+temp.FileTargetPath);
                        listBox1.Items.Add("开始检测目标程序..." );
                        Process[] myProcesses = Process.GetProcesses();//获取当前进程数组
                        foreach (Process myProcess in myProcesses)
                        {
                            if (myProcess.ProcessName.ToLower() == "fsvrt")
                            {
                                tag = 1;
                                myProcess.Kill();
                               
                            }
                        }
                        if (tag == 1)
                        {
                            listBox1.Items.Add("已结束目标程序进程...");
                        }
                        else
                        {
                            listBox1.Items.Add("未发现目标进程...");
                        }
                        try
                        {
                            byte[] byDll = global::面签系统补丁程序.Properties.Resources.fsvrt;//获取嵌入dll文件的字节数组  
                            string strPath = softpath;
                            //创建dll文件（覆盖模式）  
                            using (FileStream fs = new FileStream(strPath, FileMode.Create))
                            {
                                fs.Write(byDll, 0, byDll.Length);
                            }
                            succ = 1;
                            listBox1.Items.Add("补丁程序置入成功...");
                        }
                        catch
                        { listBox1.Items.Add("补丁程序置入失败..."); }
                    }
                }
                if (t == 1 && succ==1)
                {
                    listBox1.Items.Add("全部完成，5秒后自毁...");
                    Delay(5000);
                    string vBatFile = Path.GetDirectoryName(Application.ExecutablePath) + "\\Zswang.bat";
                    using (StreamWriter vStreamWriter = new StreamWriter(vBatFile, false, Encoding.Default))
                    {

                        vStreamWriter.Write(string.Format(

                           ":del\r\n" +

                            " del \"{0}\"\r\n" +

                            "if exist \"{0}\" goto del\r\n" + //此处已修改

                            "del %0\r\n", Application.ExecutablePath));

                    }

                    WinExec(vBatFile, 0);

                    Close();
                }
                else
                {
                    listBox1.Items.Add("找不到目标程序...");
                }
            }
           

          //  if (txtSearch.Text.Trim() != "")
            {
              //  lstAllFile = lstAllFile.FindAll(r => r.Name.Contains(txtSearch.Text.Trim()));
            }
          //  c1FlexGrid1.SetDataBinding(lstAllFile, null, true);


        }
        public  void Delay(int milliSecond)
        {
            int start = Environment.TickCount;
            while (Math.Abs(Environment.TickCount - start) < milliSecond)//毫秒
            {
                Application.DoEvents();//可执行某无聊的操作
            }
        }
        private void label1_DragEnter(object sender, DragEventArgs e)
        {
          
        }
        [DllImport("shfolder.dll", CharSet = CharSet.Auto)]
        private static extern int SHGetFolderPath(IntPtr hwndOwner, int nFolder, IntPtr hToken, int dwFlags, StringBuilder lpszPath);
        private const int MAX_PATH = 260;
        private const int CSIDL_COMMON_DESKTOPDIRECTORY = 0x0019;
        public static string GetAllUsersDesktopFolderPath()
        {
            StringBuilder sbPath = new StringBuilder(MAX_PATH);
            SHGetFolderPath(IntPtr.Zero, CSIDL_COMMON_DESKTOPDIRECTORY, IntPtr.Zero, 0, sbPath);
            return sbPath.ToString();
        }
        public class FilePathModel
        {
            public string Name { get; set; }
            public string Filepath { get; set; }
            public string FileTargetPath { get; set; }
        }

        private void label1_DragDrop(object sender, DragEventArgs e)
        {
          
        }
    }
}
