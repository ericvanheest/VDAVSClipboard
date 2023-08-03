using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace VDAVSClipboard
{
    class Program
    {
        const int WM_SYSCOMMAND = 0x0112;
        const int SC_MINIMIZE = 0xF020;

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern bool CreateProcess(string lpApplicationName,
           string lpCommandLine, IntPtr procSecurity,
           IntPtr threadSecurity, bool bInheritHandles,
           ProcessCreation dwCreationFlags, IntPtr lpEnvironment, string lpCurrentDirectory,
           [In] ref STARTUPINFO lpStartupInfo,
           out PROCESS_INFORMATION lpProcessInformation);

        [Flags]
        public enum ProcessCreation : uint
        {
            CreateBreakawayFromJob = 0x01000000,
            CreateDefaultErrorMode = 0x04000000,
            CreateNewConsole = 0x00000010,
            CreateNewProcessGroup = 0x00000200,
            CreateNoWindow = 0x08000000,
            CreateProtectedProcess = 0x00040000,
            CreatePreserveCodeAuthzLevel = 0x02000000,
            CreateSeparateWOWVDM = 0x00000800,
            CreateSharedWOWVDM = 0x00001000,
            CreateSuspended = 0x00000004,
            CreateUnicodeEnvironment = 0x00000400,
            DebugOnlyThisProcess = 0x00000002,
            DebugProcess = 0x00000001,
            DetachedProcess = 0x00000008,
            ExtendedStartupinfoPresent = 0x00080000,
        }

        public enum ShowWindowCommands : int
        {
            Hide = 0,
            Normal = 1,
            ShowMinimized = 2,
            Maximize = 3, // is this the right value?
            ShowMaximized = 3,
            ShowNoActivate = 4,
            Show = 5,
            Minimize = 6,
            ShowMinNoActive = 7,
            ShowNA = 8,
            Restore = 9,
            ShowDefault = 10,
            ForceMinimize = 11
        }

        [Flags]
        public enum StartupInfoFlags : uint
        {
            UseShowWindow = 0x00000001,
            UseSize = 0x00000002,
            UsePosition = 0x00000004,
            UseCountChars = 0x00000008,
            UseFillAttribute = 0x00000010,
            RunFullScreen = 0x00000020,
            ForceOnFeedback = 0x00000040,
            ForceOffFeedback = 0x00000080,
            UseStdHandles = 0x00000100,
            UseHotkey = 0x00000200,
            TitleIsLinkName = 0x00000800,
            TitleIsAppID = 0x00001000,
            PreventPinning = 0x00002000,
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct STARTUPINFO
        {
            public Int32 cb;
            public string lpReserved;
            public string lpDesktop;
            public string lpTitle;
            public Int32 dwX;
            public Int32 dwY;
            public Int32 dwXSize;
            public Int32 dwYSize;
            public Int32 dwXCountChars;
            public Int32 dwYCountChars;
            public Int32 dwFillAttribute;
            public StartupInfoFlags dwFlags;
            public Int16 wShowWindow;
            public Int16 cbReserved2;
            public IntPtr lpReserved2;
            public IntPtr hStdInput;
            public IntPtr hStdOutput;
            public IntPtr hStdError;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PROCESS_INFORMATION
        {
            public IntPtr hProcess;
            public IntPtr hThread;
            public int dwProcessId;
            public int dwThreadId;
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        enum AspectRatios { None, AR43, AR169 };

        public static bool bHD = false;
        public static bool bBluRayFormat = false;

        [STAThread]
        static void Main(string[] args)
        {

            AspectRatios ar = AspectRatios.None;

            if (args.Length < 1)
            {
                Console.WriteLine("Usage: VDAVSClipboard {{filename}}");
                return;
            }

            SendMessage(Process.GetCurrentProcess().MainWindowHandle, WM_SYSCOMMAND, new IntPtr(SC_MINIMIZE), IntPtr.Zero);

            string sDirectory = Path.GetDirectoryName(args[0]);
            string sFilename = Path.GetFileNameWithoutExtension(args[0]);

            if (sDirectory.EndsWith("\\Processed"))
                sDirectory = sDirectory.Substring(0, sDirectory.Length - "\\Processed".Length);

            if (sDirectory.EndsWith("\\Processing"))
                sDirectory = sDirectory.Substring(0, sDirectory.Length - "\\Processing".Length);

            int iIndex = sDirectory.LastIndexOf('\\');
            if (iIndex != -1)
                sDirectory = sDirectory.Substring(iIndex + 1);

            if (Char.IsDigit(sFilename[0]))
            {
                bBluRayFormat = true;
                bHD = true;
            }

            if (sFilename.Contains(", 480p"))
                bHD = false;
            if (sFilename.Contains(", 1080p"))
                bHD = true;
            if (sFilename.Contains(", 1080i"))
                bHD = true;

            if (sFilename.StartsWith("Video-"))
                sFilename = sFilename.Substring("Video-".Length);

            if (sFilename.EndsWith("fps"))
            {
                sFilename = Regex.Replace(sFilename, "-[0-9]*to[0-9]*fps$", "");
                sFilename = Regex.Replace(sFilename, "-[0-9]*fps$", "");
            }

            if (sFilename.EndsWith("-auto"))
            {
                sFilename = sFilename.Substring(0, sFilename.Length - "-auto".Length);
            }

            if (sFilename.EndsWith("-Unmod"))
            {
                sFilename = sFilename.Substring(0, sFilename.Length - "-unmod".Length);
            }

            if (sFilename.EndsWith("-16to9"))
            {
                ar = AspectRatios.AR169;
                sFilename = sFilename.Substring(0, sFilename.Length - "-16to9".Length);
            }
            if (sFilename.EndsWith("_(16-9)"))
            {
                ar = AspectRatios.AR169;
            }

            else if (sFilename.EndsWith("-4to3"))
            {
                ar = AspectRatios.AR43;
                sFilename = sFilename.Substring(0, sFilename.Length - "-4to3".Length);
            }
            if (sFilename.EndsWith("_(4-3)"))
            {
                ar = AspectRatios.AR169;
            }
            if (sFilename.EndsWith("_(20-11)"))
            {
                ar = AspectRatios.AR169;
            }

            if (!bBluRayFormat)
                sFilename += "-q3.avi";
            else
                sFilename += ".avi";

            string sCopyName = sDirectory + "-" + sFilename;

            Process[] processList = Process.GetProcesses();

            Process pVirtualdub = null;

            foreach(Process p in processList)
            {
                if (p.ProcessName == "send")
                {
                    p.Kill();
                    continue;
                }

                if (p.ProcessName.StartsWith("VirtualDub"))
                {
                    if (p.MainWindowTitle.StartsWith("VirtualDub"))
                    {
                        pVirtualdub = p;
                    }
                }
            }

            sCopyName = Path.Combine("F:\\Temp\\DVD", sCopyName);

            if (pVirtualdub != null)
                SendFilename(pVirtualdub, args[0], sCopyName, ar);
            else
                Console.WriteLine("Could not find unused VirtualDub process");

        }

        static string Quote(string str)
        {
            StringBuilder sb = new StringBuilder(str.Length + 6);
            int index = 0;
            while (index < str.Length)
            {
                switch (str[index])
                {
                    case '\'':
                        sb.Append("\\'");
                        break;
                    case '\\':
                        sb.Append("\\\\");
                        break;
                    default:
                        sb.Append(str[index]);
                        break;
                }
                index++;
            }
            return sb.ToString();
        }

        static void SendFilename(Process p, string sSource, string sDest, AspectRatios ar)
        {
            string sDrive = "T:";
            if (sSource[1] == ':')
                sDrive = sSource.Substring(0,2);

            string sAR43 = Quote(sDrive + @"\VD_Scripts\Main\4-3_Miscellaneous.syl");
            string sAR169 = Quote(sDrive + @"\VD_Scripts\Main\16-9_Anamorphic.syl");
            string sHDAR43 = Quote(sDrive + @"\VD_Scripts\Main\BluRay\BR-Noisy-4-3_in_16-9_Letterboxed.vcf");
            string sHDAR169 = Quote(sDrive + @"\VD_Scripts\Main\BluRay\BR-16-9_1280x720-q3.2.vcf");

            string sPause = "usleep 100";

            string sTemp = Path.GetTempFileName();

            // Create the file-open dialog
            StreamWriter writer = new StreamWriter(sTemp);
//            string sVDWindow = "\"" + p.MainWindowTitle.Replace(@"\", @"\\") + "\"";
            string sVDWindow = "\"VirtualDub \"";
//            string sSendCmd = "-wd -s VirtualDub " + sVDWindow + @" \{s100}\co\C";
            string sSendCmd = "-wb -s VirtualDub " + sVDWindow + @" ""\[WM_COMMAND,40001,0,p]""";
            writer.WriteLine("send " + sSendCmd);

            writer.WriteLine(sPause);

            // Open the sSource video
            sSendCmd = @"-wbr -i c,n,n,n,n,n,n,n,n ""^Open video file$"" ""\[WM_SETTEXT,0,'" + Quote(sSource) + "']\\[WM_COMMAND,1,0]\"";
            writer.WriteLine("send " + sSendCmd);

            sSendCmd = @"-xr ""^Open video file$""";
            writer.WriteLine("send " + sSendCmd);

            writer.WriteLine(sPause);

            string sDirectory = Path.GetDirectoryName(sSource);
            if (!File.Exists(Path.Combine(sDirectory, "New Text Document.txt")))
            {
                // Run script
                sSendCmd = "-wb -s VirtualDub " + sVDWindow + @" ""\[WM_COMMAND,40128,0,p]""";
                writer.WriteLine("send " + sSendCmd);

                writer.WriteLine(sPause);

                // Select the 4:3 or 16:9 generic script
                if (bHD)
                    sSendCmd = "-wbr -i c,n,n \"^Load configuration script$\" \"\\[WM_SETTEXT,0,'" + (ar == AspectRatios.AR43 ? sHDAR43 : sHDAR169) + "']\\[WM_COMMAND,1,0]\"";
                else
                    sSendCmd = "-wbr -i c,n,n \"^Load configuration script$\" \"\\[WM_SETTEXT,0,'" + (ar == AspectRatios.AR43 ? sAR43 : sAR169) + "']\\[WM_COMMAND,1,0]\"";
                writer.WriteLine("send " + sSendCmd);

                sSendCmd = "-xr \"^Load configuration script$\"";
                writer.WriteLine("send " + sSendCmd);

                writer.WriteLine(sPause);

                // Script workaround
                sSendCmd = "-wb -s VirtualDub " + sVDWindow + @" ""\[WM_COMMAND,40004,0,p]""";
                writer.WriteLine("send " + sSendCmd);

                sSendCmd = "-wbr \"^Filters$\" \"\\[WM_COMMAND,1,0]\"";
                writer.WriteLine("send " + sSendCmd);

                sSendCmd = "-xr \"^Filters$\"";
                writer.WriteLine("send " + sSendCmd);
            }

            writer.WriteLine(sPause);

            //sSendCmd = "-sd VirtualDub " + sVDWindow + " 1";
            //writer.WriteLine("send " + sSendCmd);

            string sTarget = sDest.Replace(".h264.dga.", ".");
            sTarget = sTarget.Replace(".vc1.dga.", ".");
            sTarget = sTarget.Replace(".mkv.dga.", ".");

            // Wait for save file dialog
//            sSendCmd = @"-wd ""Save AVI 2.0 File"" ""\{s100}" + sDest.Replace(@"\", @"\\");
            sSendCmd = @"-wbr -i c,c,c,c,c ""^Save AVI 2\.0 File$"" ""\[WM_SETTEXT,0,'" + Quote(sTarget) + "']\"";
            writer.WriteLine("send " + sSendCmd);

            writer.Close();
            File.Move(sTemp, sTemp + ".cmd");

            string sCmd = Environment.GetFolderPath(Environment.SpecialFolder.System) + "\\cmd.exe";
//            Console.WriteLine("Executing: " + sCmd + "/C \"" + sTemp + ".cmd\"");
//            Console.ReadLine();

            PROCESS_INFORMATION pi;
            pi.dwProcessId = 0;
            pi.dwThreadId = 0;
            pi.hProcess = IntPtr.Zero;
            pi.hThread = IntPtr.Zero;
            STARTUPINFO si = new STARTUPINFO();
            si.cb = Marshal.SizeOf(si);
            si.dwFlags = StartupInfoFlags.UseShowWindow;
            si.cbReserved2 = 0;
            si.lpReserved = null;
            si.lpReserved2 = IntPtr.Zero;
            si.lpDesktop = null;
            si.wShowWindow = (short)ShowWindowCommands.Minimize;
//            Process pTemp = Process.Start(sCmd, "/C \"" + sTemp + ".cmd\"");
            if (CreateProcess(sCmd, sCmd + "/C \"" + sTemp + ".cmd\"", IntPtr.Zero, IntPtr.Zero, false, 0, IntPtr.Zero, null, ref si, out pi))
            {
                Process pTemp = Process.GetProcessById(pi.dwProcessId);
                pTemp.WaitForExit();
            }
            else
            {
                Console.WriteLine("Unable to create process: {0}", sCmd);
            }

            File.Delete(sTemp + ".cmd");
        }
    }
}
