using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Timers;
using System.Runtime.InteropServices;
using System.IO;
using AccessDLL;
using DrawCurve;
using XMLConfigPLC200SampleData;

namespace KingStoneModify
{
    public partial class Bench_3 : Form
    {
        #region INI
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);
        #endregion

        #region AccessDll
        [DllImport("AccessDLL.dll", EntryPoint = "Init", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void Init(string FileName);

        [DllImport("AccessDLL.dll", EntryPoint = "CreatFile", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void CreatFile();

        [DllImport("AccessDLL.dll", EntryPoint = "AddTable", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void AddTable(string TableName);

        [DllImport("AccessDLL.dll", EntryPoint = "AddColumn", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void AddColumn(string TableName, string ColName, string Datatype);

        [DllImport("AccessDLL.dll", EntryPoint = "AddNewRow", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void AddNewRow(string TableName);

        [DllImport("AccessDLL.dll", EntryPoint = "DeleteRow", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void DeleteRow(string TableName, int index);

        [DllImport("AccessDLL.dll", EntryPoint = "WriteFloatsData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteFloatsData(string TableName, int index, string ColName, float[] data);

        [DllImport("AccessDLL.dll", EntryPoint = "WriteIntsData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteIntsData(string TableName, int index, string ColName, int[] data);

        [DllImport("AccessDLL.dll", EntryPoint = "WriteSigleData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteSigleData(string TableName, int index, string ColName, object data);

        [DllImport("AccessDLL.dll", EntryPoint = "ReadFloatsData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ReadFloatsData(string TableName, int index, string ColName, out float[] data);

        [DllImport("AccessDLL.dll", EntryPoint = "ReadIntsData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ReadIntsData(string TableName, int index, string ColName, out int[] data);

        [DllImport("AccessDLL.dll", EntryPoint = "ReadSigleData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ReadSigleData(string TableName, int index, string ColName, ref object data);

        [DllImport("AccessDLL.dll", EntryPoint = "QueryData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void QueryData(string TableName, int index, string ColName, ref object data);

        [DllImport("AccessDLL.dll", EntryPoint = "Init_M", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void Init_M(string FileName, string TableName);

        [DllImport("AccessDLL.dll", EntryPoint = "WriteFloatsData_M", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteFloatsData_M(string TableName, int index, string ColName, float[] data);

        [DllImport("AccessDLL.dll", EntryPoint = "WriteIntsData_M", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteIntsData_M(string TableName, int index, string ColName, int[] data);

        [DllImport("AccessDLL.dll", EntryPoint = "WriteSigleData_M", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteSigleData_M(string TableName, int index, string ColName, object data);

        [DllImport("AccessDLL.dll", EntryPoint = "SaveDate_M", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void SaveDate_M();

        [DllImport("AccessDLL.dll", EntryPoint = "DeleteRow_M", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void DeleteRow_M(string TableName, int index);

        [DllImport("AccessDLL.dll", EntryPoint = "ReadFloatsData_M", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ReadFloatsData_M(string TableName, int index, string ColName, out float[] data);

        [DllImport("AccessDLL.dll", EntryPoint = "ReadIntsData_M", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ReadIntsData_M(string TableName, int index, string ColName, out int[] data);

        [DllImport("AccessDLL.dll", EntryPoint = "ReadSigleData_M", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ReadSigleData_M(string TableName, int index, string ColName, ref object data);

        [DllImport("AccessDLL.dll", EntryPoint = "QueryData_M", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void QueryData_M(string TableName, int index, string ColName, ref object data);
        #endregion

        #region DrawCurve
        [DllImport("DrawCurve.dll", EntryPoint = "Init", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void Init(Panel p, string CurveTypee);

        [DllImport("DrawCurve.dll", EntryPoint = "AddPointToList", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void AddPointToList(PointF[] p);

        [DllImport("DrawCurve.dll", EntryPoint = "DrawModule", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void DrawModule(Panel panel);

        [DllImport("DrawCurve.dll", EntryPoint = "DeletePointFromList", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void DeletePointFromList(int index);

        [DllImport("DrawCurve.dll", EntryPoint = "GetZoomPara", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void GetZoomPara(ref float x, ref float y);

        [DllImport("DrawCurve.dll", EntryPoint = "SetZoomPara", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void SetZoomPara(float x, float y);

        [DllImport("DrawCurve.dll", EntryPoint = "DrawLine", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void DrawLine(PointF[] p, Panel panel, Color col);

        [DllImport("DrawCurve.dll", EntryPoint = "DrawLine", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void DrawLine(PointF[] p, Panel panel, Color col, float xMax, float yMax);

        [DllImport("DrawCurve.dll", EntryPoint = "DrawMultiLine", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void DrawMultiLine(Panel panel, Color[] col);

        [DllImport("DrawCurve.dll", EntryPoint = "DrawMultiLine", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void DrawMultiLine(Panel panel, Color[] col, float xMax, float yMax);

        [DllImport("DrawCurve.dll", EntryPoint = "DrawMultiLine", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void GetZoomAreaPoint(PointF[] p, out PointF[] f, Panel panel, PointF StartP, PointF EndP);

        #endregion

        private string strFilePath = Application.StartupPath + @"\3#\TestConfig.ini";//获取INI文件路径
        private string strSec = "TestConfig"; //INI文件名

        private Form1 m_MainFrameHandle;
        private GroupBox m_GrpParent;
        private SettingForm_No3 m_SettingFormHandle;
        private TestResultShow_No3 m_TestResultHandle;
        private CreateTestResult_No3 m_CreateTestResult_No3Handle;
        public AccessDLLInterface m_AccessHandle = new AccessDLLInterface();
        public DrawCurve.Interface m_DrawCurveHandle = new Interface();
        public string m_LineType = "Line";//选择线型
        public UInt16 m_SetTestTimes = 0;
        public UInt32 m_yMax = 250;
        public Bench_3(Form1 Handle, GroupBox grp)
        {
            InitializeComponent();
            Bench_3.CheckForIllegalCrossThreadCalls = false;
            m_MainFrameHandle = Handle;
            m_GrpParent = grp;
        }

        private System.Timers.Timer m_SampleDataTimer = new System.Timers.Timer();
        private System.Timers.Timer m_DrawLineTimer = new System.Timers.Timer();
        private System.Timers.Timer m_SampleSystemStatusTimer = new System.Timers.Timer();
        private System.Timers.Timer m_SampleStartTimer = new System.Timers.Timer();
        private System.Timers.Timer m_SampleStopTimer = new System.Timers.Timer();
        public string m_DataBaseFilePath = Application.StartupPath + "\\3#\\DataBase.mdb";

        public string m_TestNo = default(string);
        public int m_TestSequence = 1;//试验顺序；

        public class PointFArray
        {
            public PointF m_Pt;
            public PointFArray(PointF pt)
            {
                m_Pt = pt;
            }
        }
        public List<PointFArray> m_PointFArrays = new List<PointFArray>();
        private object O_LockPointFArray = new object();
        private bool m_isFirtSampleDate = true;

        public class PointFList
        {
            public PointF[] m_pts;
            public PointFList(PointF[] pt)
            {
                m_pts = new PointF[pt.Length];
                for (int i = 0; i < pt.Length; i++)
                {
                    m_pts[i] = pt[i];
                }
            }
        }
        public List<PointFList> m_PointFLists = new List<PointFList>();

        public class TestResultList
        {
            public string m_No;
            public float m_InitPressure;
            public float m_EndPressure;
            public float m_KeepTime;
            public float m_DropPressure;
            public TestResultList(string No, float IPre, float EPre, float KTime, float DPre)
            {
                m_No = No;
                m_InitPressure = IPre;
                m_EndPressure = EPre;
                m_KeepTime = KTime;
                m_DropPressure = DPre;
            }
        }
        public List<TestResultList> m_TestResultLists = new List<TestResultList>();
        private void Bench_3_Load(object sender, EventArgs e)
        {
            m_SettingFormHandle = new SettingForm_No3(m_MainFrameHandle, this);
            m_SettingFormHandle.TopLevel = false;
            m_GrpParent.Controls.Add(m_SettingFormHandle);

            m_SampleDataTimer.AutoReset = true;
            m_SampleDataTimer.Interval = 500;
            m_SampleDataTimer.Elapsed += new ElapsedEventHandler(SampleDataFun);

            m_DrawLineTimer.AutoReset = true;
            m_DrawLineTimer.Interval = 1000;
            m_DrawLineTimer.Elapsed += new ElapsedEventHandler(DrawLineFun);

            m_SampleSystemStatusTimer.AutoReset = true;
            m_SampleSystemStatusTimer.Interval = 500;
            m_SampleSystemStatusTimer.Elapsed += new ElapsedEventHandler(SampleSystemStatusFun);
            m_SampleSystemStatusTimer.Enabled = true;


            m_SampleStartTimer.AutoReset = true;
            m_SampleStartTimer.Interval = 500;
            m_SampleStartTimer.Elapsed += new ElapsedEventHandler(SampleStartFun);
            m_SampleStartTimer.Enabled = true;


            m_SampleStopTimer.AutoReset = true;
            m_SampleStopTimer.Interval = 500;
            m_SampleStopTimer.Elapsed += new ElapsedEventHandler(SampleStopFun);

            string strSec = Path.GetFileNameWithoutExtension(strFilePath);
            m_TestNo = ContentValue(strSec, "TestNo_Bench3");

            m_DrawCurveHandle.Init(panel_No1, m_LineType);

            AccessInit();

            BT_Start.Enabled = false;

            textBox_yMax.Text = m_yMax.ToString();
            m_DrawCurveHandle.DrawModule(panel_No1);
        }

        private string ContentValue(string Section, string key)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }

        /// <summary>
        /// 数据库初始化
        /// </summary>
        private void AccessInit()
        {
            m_AccessHandle.Init(m_DataBaseFilePath);
            m_AccessHandle.Init_M(m_DataBaseFilePath, "TestResult_1");
        }

        /// <summary>
        /// 采集PLC压力曲线数据
        /// </summary>
        private bool m_isStartSample = true;
        private void SampleDataFun(object o, ElapsedEventArgs e)
        {
            float TestData = 0;
            DateTime dt = DateTime.Now;

            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("PressureTest_Bench3", ref TestData, ref dt);
            if (code != 1)
            {
                return;
            }

            if (m_isFirtSampleDate)
            {
                m_BaseTime = dt;
                m_isFirtSampleDate = false;
                TestData = 0.3f;
                m_DrawCurveTimeSpan = 30;
            }
            if (m_isStartSample)
            {
                TestData = 0.3f;
                m_DrawCurveTimeSpan = 30;
                m_isStartSample = false;
            }

            if (TestData <= 0.3)
            {
                TestData = 0.3f;
            }

            float TotalTime = 0;
            GetTimeSpan(dt, m_BaseTime, ref TotalTime);
            PointF temp = new PointF(0, 0);
            temp.X = TotalTime;
            temp.Y = TestData;
            lock (O_LockPointFArray)
            {
                m_PointFArrays.Add(new PointFArray(temp));
            }
            lock (o_EndTimeLock)
            {
                m_EndTime = dt;
            }

            float KeepPressureTime = 0.0f;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressureTime_Bench3", ref KeepPressureTime, ref dt);
            if (code != 1)
            {
                return;
            }
            TextBox_KeepPressureTime.Text = KeepPressureTime.ToString("0.00");

            float DropPressure = 0.0f;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("DropPressure_Bench3", ref DropPressure, ref dt);
            if (code != 1)
            {
                return;
            }
            TextBox_KeepPressureDrop.Text = DropPressure.ToString("0.00");
        }

        public void ReadEndTime(ref DateTime EndTime)
        {
            lock (o_EndTimeLock)
            {
                EndTime = m_EndTime;
            }
        }

        /// <summary>
        /// 采集系统状态
        /// </summary>
        public DateTime m_BaseTime;
        public bool m_isSetFlag = false;
        private void SampleSystemStatusFun(object o, ElapsedEventArgs e)
        {
            bool isCommunicationFalt = false;
            m_MainFrameHandle.m_PLCCommHandle.ReportComState(ref isCommunicationFalt);
            if (isCommunicationFalt)
            {
                TextBox_Status.Text = "通信故障";
                TextBox_Status.BackColor = Color.Red;
                return;
            }
            else
            {
                TextBox_Status.BackColor = Color.Black;
            }

            bool SystemStatusManual = false;
            DateTime dt = DateTime.Now;
            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusManual_No1_Bench3", ref SystemStatusManual, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusManual)
            {
                TextBox_Status.Text = "手动";
            }

            bool SystemStatusEmergencyStop = false;
            dt = DateTime.Now;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusEmergencyStop_No1_Bench3", ref SystemStatusEmergencyStop, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusEmergencyStop)
            {
                TextBox_Status.Text = "急停";
            }

            bool SystemStatusStop = false;
            dt = DateTime.Now;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusStop_No1_Bench3", ref SystemStatusStop, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusStop)
            {
                TextBox_Status.Text = "停止";
                BT_Start.Enabled = true;
            }

            bool SystemStatusAddPressure = false;
            dt = DateTime.Now;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusAddPressure_No1_Bench3", ref SystemStatusAddPressure, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusAddPressure)
            {
                TextBox_Status.Text = "加压";
            }

            bool SystemStatusKeepPressure = false;
            dt = DateTime.Now;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusKeepPressure_No1_Bench3", ref SystemStatusKeepPressure, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusKeepPressure)
            {
                TextBox_Status.Text = "保压";
            }

            bool SystemStatusDropPressure = false;
            dt = DateTime.Now;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusDropPressure_No1_Bench3", ref SystemStatusDropPressure, ref dt);
            if (code != 1)
            {
                return;
            }

            float PressureTest = 0.0f;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("PressureTest_Bench3", ref PressureTest, ref dt);
            if (code != 1)
            {
                return;
            }
            TextBox_TestPressure.Text = PressureTest.ToString("0.00");
            if (SystemStatusDropPressure)
            {
                TextBox_Status.Text = "泄压";
            }

            if (m_isSetFlag)
            {
                BT_Start.Enabled = true;
                m_isSetFlag = false;
            }

            //ReadSettingInfo();
            //ReadChanelInfo();

        }

        private void GetTimeSpan(DateTime now, DateTime BaseTime, ref float TotalTime)
        {
            TimeSpan tp = now - BaseTime;
            TotalTime = (tp.Days * 24 * 3600) +
                        (tp.Hours * 3600) +
                        (tp.Minutes * 60) +
                        (tp.Seconds) +
                        tp.Milliseconds / 1000.0f;
        }

        public bool m_isTesting = false;//是否正在测试
        private void SampleStartFun(object o, ElapsedEventArgs e)
        {
            bool startFlag = false;
            DateTime dt = DateTime.Now;
            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ReadStartTest_Bench3", ref startFlag, ref dt);
            if (code != 1)
            {
                return;
            }

            if (startFlag)
            {
                //ReadSettingData();
                //SaveSettingData();

                //m_PointFArrays.Clear();
                m_SampleDataTimer.Start();
                //for (int i = 0; i < 300000000; i++) ;
                m_DrawLineTimer.Start();
                m_SampleStartTimer.Stop();
                m_SampleStopTimer.Start();

                if (m_isTestStartBaseTestNo)
                {
                    m_StartTime = DateTime.Now;
                    m_isTestStartBaseTestNo = false;
                }

                label_X1.Text = m_BaseTime.ToString("HH:mm:ss");
                DateTime d1 = m_BaseTime;
                d1.AddMinutes(m_DrawCurveTimeSpan / 60);
                label_X7.Text = d1.ToString("HH:mm:ss");
                BT_Start.Enabled = false;
                m_isTesting = true;
            }

            startFlag = false;
            code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ReadStartTest_Bench3", startFlag);
        }

        private void SampleStopFun(object o, ElapsedEventArgs e)
        {
            //采集PLC停止命令
            bool stopFlag = false;
            DateTime dt = DateTime.Now;
            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ReadStopTest_Bench3", ref stopFlag, ref dt);
            if (code != 1)
            {
                return;
            }

            if (stopFlag)
            {
                m_DrawLineTimer.Stop();
                m_SampleDataTimer.Stop();
                m_SampleStopTimer.Stop();
                m_SampleStartTimer.Start();
                m_isFirtSampleDate = true;
                m_isStartSample = true;
                m_isTesting = false;
                //label_X2.Text = DateTime.Now.ToString("HH:mm:ss");
                //CreatResultForm();
                //ShowFun();    

            }
            bool stopFlag1 = false;
            code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ReadStopTest_Bench3", stopFlag1);


            if (stopFlag)
            {
                BT_Start.Enabled = true;
                MethodInvoker invoke = new MethodInvoker(CreatResultForm);
                BeginInvoke(invoke);
            }
        }

        public int m_DrawCurveTimeSpan = 30;//画曲线时间宽度,初始化值为30s
        private void DrawLineFun(object o, ElapsedEventArgs e)
        {
            int len = 0;
            lock (O_LockPointFArray)
            {
                len = m_PointFArrays.Count;
            }
            if (len < 4)
            {
                return;
            }
            PointF[] p = new PointF[len];
            lock (O_LockPointFArray)
            {
                for (int i = 0; i < len; i++)
                {
                    p[i] = m_PointFArrays[i].m_Pt;
                }
            }
            if (m_DrawCurveTimeSpan < 5)
            {
                m_DrawCurveTimeSpan = 10;
            }
            m_DrawCurveHandle.DrawLine(p, panel_No1, Color.Red, m_DrawCurveTimeSpan, m_yMax);
            if ((p[len - 1].X - p[0].X) > m_DrawCurveTimeSpan)
            {
                m_DrawCurveTimeSpan += 30;//如果超过，则自加30s宽度；
            }

            label_X1.Text = m_BaseTime.ToString("HH:mm:ss");

            int xMax = m_DrawCurveTimeSpan;
            int H = 0;
            int M = 0;
            int S = 0;

            GetlableTime(xMax / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(xMax * 2 / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(xMax * 3 / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(xMax * 4 / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(xMax * 5 / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(xMax * 6 / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

        }

        private void GetlableTime(int t, ref int H, ref int M, ref int S)
        {
            H = (int)(t / 3600);
            M = (int)((t - H * 3600) / 60);
            S = (int)((t - H * 3600) % 60);
        }

        private void BT_Setting_Click(object sender, EventArgs e)
        {
            if (m_isTesting)
            {
                MessageBox.Show("请先停止测试。", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            this.Hide();
            m_SettingFormHandle.Show();
        }

        public DateTime m_StartTime;//本轮试验起始时间
        public bool m_isTestStartBaseTestNo = true;//基于试验编号的试验起始标志
        public DateTime m_EndTime;//本轮试压结束时间
        object o_EndTimeLock = new object();
        private void BT_Start_Click(object sender, EventArgs e)
        {
            bool StartSignal = true;
            m_SampleDataTimer.Start();
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("StartTest_Bench3", StartSignal);
            if (code != 1)
            {
                MessageBox.Show("启动失败，请重新启动", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            TextBox_Status.Text = "开始";
            m_SampleDataTimer.Enabled = true;
            //m_StartTime = DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss");    

            //m_PointFArrays.Clear();//清采集数据

            //m_DrawLineTimer.Start();
            //m_SampleDataTimer.Start();
            //m_SampleStopTimer.Start();           
            //m_SampleStartTimer.Stop();


            if (m_isTestStartBaseTestNo)
            {
                m_StartTime = DateTime.Now;
                m_isTestStartBaseTestNo = false;
            }

            label_X1.Text = m_BaseTime.ToString("HH:mm:ss");
            DateTime d1 = m_BaseTime;
            //d1.AddMinutes(m_DrawCurveTimeSpan / 60);
            //label_X4.Text = d1.ToString("HH:mm:ss");

            BT_Start.Enabled = false;
        }

        private void BT_Stop_Click(object sender, EventArgs e)
        {
            bool StopSignal = true;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("StopTest_Bench3", StopSignal);
            if (code != 1)
            {
                MessageBox.Show("停止失败，请重新启动", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //m_isFirtSampleDate = true;

            //m_DrawLineTimer.Stop();
            //m_SampleDataTimer.Stop();
            //m_SampleStopTimer.Stop();
            //m_SampleStartTimer.Start();

            //CreatResultForm();

            //label_X2.Text = DateTime.Now.ToString("HH:mm:ss");
            //label_X2.Visible = false;

            //BT_Start.Enabled = true;
        }

        /// <summary>
        /// 将点加入到List列表
        /// </summary>
        public void AddPonitFtoList()
        {
            int len = 0;
            lock (O_LockPointFArray)
            {
                len = m_PointFArrays.Count;
            }
            if (len <= 1)
            {
                return;
            }
            PointF[] temp = new PointF[len];
            TimeSpan ts = m_BaseTime - m_StartTime;
            float s = ts.Hours * 3600 + ts.Minutes * 60 + ts.Seconds;
            lock (O_LockPointFArray)
            {
                for (int i = 0; i < len; i++)
                {
                    if (i == 0)
                    {
                        temp[i].X = s;
                        temp[i].Y = m_PointFArrays[i].m_Pt.Y;
                        continue;
                    }
                    temp[i] = m_PointFArrays[i].m_Pt;
                }
            }

            m_PointFLists.Add(new PointFList(temp));
        }
        /// <summary>
        /// 清List表
        /// </summary>
        public void ClearPointFList()
        {
            m_PointFLists.Clear();
        }

        private void BT_Result_Click(object sender, EventArgs e)
        {
            CreatResultForm();
            //this.Hide(); 
        }

        public void CreatResultForm()
        {
            if (m_TestResultHandle != null)
            {
                m_TestResultHandle = null;
            }
            m_TestResultHandle = new TestResultShow_No3(this, m_MainFrameHandle);
            m_TestResultHandle.TopLevel = false;
            m_GrpParent.Controls.Add(m_TestResultHandle);
            m_TestResultHandle.Show();
            this.Hide();
        }

        private void BT_Report_Click(object sender, EventArgs e)
        {
            m_CreateTestResult_No3Handle = new CreateTestResult_No3(this);
            m_CreateTestResult_No3Handle.ShowDialog();

            m_CreateTestResult_No3Handle.Close();
            m_CreateTestResult_No3Handle.Dispose();
        }

        private void ReadSettingData()
        {
            bool ChanelSelect_No1 = false;
            DateTime dt = default(DateTime);
            m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No1_Bench3", ref ChanelSelect_No1, ref dt);

            bool ChanelSelect_No2 = false;
            m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No2_Bench3", ref ChanelSelect_No2, ref dt);

            bool ChanelSelect_No3 = false;
            m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No3_Bench3", ref ChanelSelect_No3, ref dt);

            float ContinueTime = 0.0f;
            if (ChanelSelect_No3)
            {
                float time3 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No3_Bench3", ref time3, ref dt);
                ContinueTime += time3;

                float Press3 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No3_Bench3", ref Press3, ref dt);

                WritePrivateProfileString(strSec, "KeepPressure_No3_Bench3", Press3.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepTime_No3_Bench3", time3.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepSelect_No3_Bench3", "ON", strFilePath);
            }

            if (ChanelSelect_No2)
            {
                float time2 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No2_Bench3", ref time2, ref dt);
                ContinueTime += time2;

                float Press2 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No2_Bench3", ref Press2, ref dt);

                WritePrivateProfileString(strSec, "KeepPressure_No2_Bench3", Press2.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepTime_No2_Bench3", time2.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepSelect_No2_Bench3", "ON", strFilePath);
            }

            if (ChanelSelect_No1)
            {
                float time1 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No1_Bench3", ref time1, ref dt);
                ContinueTime += time1;

                float Press1 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No1_Bench3", ref Press1, ref dt);

                WritePrivateProfileString(strSec, "KeepPressure_No1_Bench3", Press1.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepTime_No1_Bench3", time1.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepSelect_No1_Bench3", "ON", strFilePath);
            }

            m_DrawCurveTimeSpan = (int)(ContinueTime * 60);
        }

        private void textBox_yMax_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m_yMax = Convert.ToUInt32(textBox_yMax.Text);
            }
            catch (Exception)
            {
                //MessageBox.Show("纵轴最大值输入错误，请核对", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                m_yMax = 250;
                return;
            }
            if (m_yMax < 1)
            {
                m_yMax = 250;
            }
            UInt32 t1 = m_yMax * 1 / 5;
            label9.Text = t1.ToString() + "MPa";

            UInt32 t2 = m_yMax * 2 / 5;
            label8.Text = t2.ToString() + "MPa";

            UInt32 t3 = m_yMax * 3 / 5;
            label7.Text = t3.ToString() + "MPa";

            UInt32 t4 = m_yMax * 4 / 5;
            label4.Text = t4.ToString() + "MPa";
        }
    }
}
