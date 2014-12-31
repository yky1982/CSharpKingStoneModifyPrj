using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Runtime.InteropServices;
using System.IO;
using DrawCurve;
using System.Timers;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System.Reflection;

namespace KingStoneModify
{
    public partial class CreateTestResult_No1 : Form
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);

        #region DrawCurve
        [DllImport("DrawCurve.dll", EntryPoint = "Init", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void Init(Panel p, string CurveTypee);

        [DllImport("DrawCurve.dll", EntryPoint = "AddPointToList", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void AddPointToList(PointF[] p);

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

        private Bench_1 m_BenchHandle;
        private string I_ProductName;
        private string I_ProductNo;
        private string I_ProductModel;
        private string I_TestStandard;
        private string I_TestNiuju;
        private string I_EquipmentNo;
        private string I_CompanyName;
        private string I_Operator;
        private string I_Tester;
        private string I_Confirmation;
        private string I_Date;
        private string m_TestNo;
        public CreateTestResult_No1(Bench_1 handle)
        {
            InitializeComponent();
            m_BenchHandle = handle;
        }

        private UInt32 m_yMax = 0;
        private float m_xMax = 0.0f;
        private System.Timers.Timer m_DrawCurve = new System.Timers.Timer();
        private System.Timers.Timer m_DrawCurve_e = new System.Timers.Timer();
        private System.Timers.Timer m_DrawHistoryCurveTimer = new System.Timers.Timer();
        private System.Timers.Timer m_DrawHistoryCurveTimer_e = new System.Timers.Timer();
        private void CreateTestResult_No1_Load(object sender, EventArgs e)
        {
            m_DrawCurve.AutoReset = false;
            m_DrawCurve.Interval = 1000;
            m_DrawCurve.Elapsed += new ElapsedEventHandler(DrawCurveFun);
            m_DrawCurve.Start();

            m_DrawCurve_e.AutoReset = false;
            m_DrawCurve_e.Interval = 1000;
            m_DrawCurve_e.Elapsed += new ElapsedEventHandler(DrawCurveFun_e);
            //m_DrawCurve.Start();

            m_DrawHistoryCurveTimer.AutoReset = false;
            m_DrawHistoryCurveTimer.Interval = 1000;
            m_DrawHistoryCurveTimer.Elapsed += new ElapsedEventHandler(DrawHistoryCurveFun);
            //m_DrawHistoryCurveTimer.Start();

            m_DrawHistoryCurveTimer_e.AutoReset = false;
            m_DrawHistoryCurveTimer_e.Interval = 1000;
            m_DrawHistoryCurveTimer_e.Elapsed += new ElapsedEventHandler(DrawHistoryCurveFun_e);
            //m_DrawHistoryCurveTimer.Start();

            m_BenchHandle.m_DrawCurveHandle.Init(panel_Result, m_BenchHandle.m_LineType);

            ReadTestInfo();
            ReadTestDate();
            ReadOfficeInfo();
            //ReadTestPic();

            m_TestNo = m_BenchHandle.m_TestNo;
            I_Date = DateTime.Now.ToString("yyyy-MM-dd");
            TB_Date.Text = I_Date;
            TB_Date_e.Text = I_Date;

            TB_Result.Text = "在保压周期内无可见渗漏";
            TB_Result_e.Text = "no visibale leakage during each holding period";

            this.ControlBox = false;

            ///调整数据单元格显示格式
            dataGridView_DateList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView_DateList_e.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;


            m_yMax = m_BenchHandle.m_yMax;
            textBox_yMax.Text = m_yMax.ToString();
            textBox_yMax_e.Text = m_yMax.ToString();
        }

        #region 中文区
        private void DrawCurveFun(object o, ElapsedEventArgs e)
        {
            ReadTestPic1();
            SavePic();
            //ReadTextBoxData();
        }

        private void DrawHistoryCurveFun(object o, ElapsedEventArgs e)
        {
            DrawHistoryCurve();
            SavePic();
        }

        /// <summary>
        /// 读取曲线，分段显示
        /// </summary>
        private PointF[] m_ChangedPointF;

        /// <summary>
        /// 将内存数据绘制成报表中的数据
        /// </summary>
        private void ReadTestPic1()
        {
            //绘制底层模板
            m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result);
            //m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result_e);

            //
            Graphics gfs = panel_Result.CreateGraphics();
            Pen mypen = new Pen(Color.Red, 1.5f);
            Image image = new Bitmap(panel_Result.Width, panel_Result.Height);
            Graphics g = Graphics.FromImage(image);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;


            //去横轴最大值
            int count = m_BenchHandle.m_PointFLists.Count;
            if (count <= 0)
            {
                return;
            }
            float xMax = 0;

            int PointCount = 0;
            for (int i = 0; i < count; i++)
            {
                int len = m_BenchHandle.m_PointFLists[i].m_pts.Length;
                //if (len <= 0)
                //{
                //    return;
                //}
                //xMax += m_BenchHandle.m_PointFLists[i].m_pts[len - 1].X + 10;
                PointCount += len;
            }
            //xMax += 2;


            int length = m_BenchHandle.m_PointFLists[count - 1].m_pts.Length;
            if (length <= 0)
            {
                return;
            }
            xMax = m_BenchHandle.m_PointFLists[count - 1].m_pts[0].X + m_BenchHandle.m_PointFLists[count - 1].m_pts[length - 1].X;
            xMax += 2;
            m_xMax = xMax;
            //float position = 0;

            m_ChangedPointF = new PointF[PointCount + count];
            int ChangedPontfCount = 0;
            //坐标变换
            for (int i = 0; i < count; i++)
            {
                int len = m_BenchHandle.m_PointFLists[i].m_pts.Length;
                if (len <= 0)
                {
                    continue;
                }
                PointF[] temp = new PointF[len];
                PointF[] Ctemp;

                for (int j = 0; j < len; j++)
                {
                    if (j == 0)
                    {
                        temp[j].X = m_BenchHandle.m_PointFLists[i].m_pts[0].X;
                        temp[j].Y = m_BenchHandle.m_PointFLists[i].m_pts[j].Y;
                    }
                    else
                    {
                        temp[j].X = m_BenchHandle.m_PointFLists[i].m_pts[0].X + m_BenchHandle.m_PointFLists[i].m_pts[j].X;
                        temp[j].Y = m_BenchHandle.m_PointFLists[i].m_pts[j].Y;
                    }
                }

                PointTranslate(temp, out Ctemp, panel_Result, xMax, m_yMax);
                for (int m = 1; m < Ctemp.Length; m++)
                {
                    g.DrawLine(mypen, Ctemp[m - 1], Ctemp[m]);
                }


                //gfs.DrawLines(mypen, Ctemp);

                //position = temp[len - 1].X;

                for (int m = 0; m < Ctemp.Length; m++)
                {
                    m_ChangedPointF[ChangedPontfCount + m] = temp[m];
                }
                ChangedPontfCount += len;
                m_ChangedPointF[ChangedPontfCount].X = 0;
                m_ChangedPointF[ChangedPontfCount].Y = 0;
                ChangedPontfCount++;
            }
            gfs.DrawImage(image, 0, 0, panel_Result.Width, panel_Result.Height);

            DateTime StartTime = m_BenchHandle.m_StartTime;
            DateTime EndTime = default(DateTime);
            m_BenchHandle.ReadEndTime(ref EndTime);
            TimeSpan tp = EndTime - StartTime;
            int second = (int)tp.TotalSeconds;
            label_X1.Text = StartTime.ToString("HH:mm:ss");
            textBox_StartTime.Text = label_X1.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7.Text = EndTime.ToString("HH:mm:ss");
            textBox_EndTime.Text = label_X7.Text;

            UInt32 t1 = m_yMax * 1 / 5;
            label32.Text = t1.ToString() + "MPa";

            UInt32 t2 = m_yMax * 2 / 5;
            label33.Text = t2.ToString() + "MPa";

            UInt32 t3 = m_yMax * 3 / 5;
            label34.Text = t3.ToString() + "MPa";

            UInt32 t4 = m_yMax * 4 / 5;
            label36.Text = t4.ToString() + "MPa";
        }

        private void GetlableTime(int t, ref int H, ref int M, ref int S)
        {
            H = (int)(t / 3600);
            M = (int)((t - H * 3600) / 60);
            S = (int)((t - H * 3600) % 60);
        }

        /// <summary>
        /// 截取图片
        /// </summary>
        /// <returns></returns>
        private bool SavePic()
        {
            //Rectangle rect = Screen.GetWorkingArea();           

            //int width = 727;
            //int height = 280;
            //Bitmap image = new Bitmap(width, height);

            //Graphics g = Graphics.FromImage(image);
            //Point p1 = default(Point);
            //Point p2 = default(Point);
            //p1.X = 0;
            //p1.Y = 0;
            //p2.X = this.panel_Result.Location.X - 30;
            //p2.Y = this.panel_Result.Location.Y + 135;

            //g.CopyFromScreen(p2, p1, image.Size);

            //string FileName = m_BaseFilePath + "\\Presure.png";
            //image.Save(FileName);

            int width = 630;
            int height = 335;
            Bitmap image = new Bitmap(width, height);

            Graphics g = Graphics.FromImage(image);
            Point p1 = default(Point);
            Point p2 = default(Point);
            p1.X = 0;
            p1.Y = 0;
            p2.X = this.Left + 58;
            p2.Y = this.Top + 337;

            g.CopyFromScreen(p2, p1, image.Size);

            string FileName = m_BaseFilePath + "\\Presure.png";
            image.Save(FileName);


            //System.Drawing.Rectangle rect = new System.Drawing.Rectangle(0, 0, panel_Result.Width, panel_Result.Height);
            //Bitmap bmp = new Bitmap(rect.Width, rect.Height);
            //panel_Result.DrawToBitmap(bmp, rect);
            //bmp.Save(@"d:\123.jpg");

            return true;
        }

        /// <summary>
        /// 绘制历史曲线
        /// </summary>
        private void DrawHistoryCurve()
        {
            m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result);
            //m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result_e);

            Graphics gfs = panel_Result.CreateGraphics();
            Pen mypen = new Pen(Color.Red, 1.5f);
            Image image = new Bitmap(panel_Result.Width, panel_Result.Height);
            Graphics g = Graphics.FromImage(image);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            //Graphics gfs_e = panel_Result_e.CreateGraphics();
            //Pen mypen_e = new Pen(Color.Red, 1.5f);
            //g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            int len = m_HistoryPointF.Length;

            int pos = 0;
            for (int i = 0; i < len; i++)
            {
                if (i != 0 && m_HistoryPointF[i].X == 0 && m_HistoryPointF[i].Y == 0)
                {
                    PointF[] f = new PointF[i - pos];
                    for (int j = pos; j < i; j++)
                    {
                        f[j - pos].X = m_HistoryPointF[j].X;
                        f[j - pos].Y = m_HistoryPointF[j].Y;
                    }
                    PointF[] cf;
                    PointTranslate(f, out cf, panel_Result, m_xMax, m_yMax);
                    for (int m = 1; m < cf.Length; m++)
                    {
                        g.DrawLine(mypen, cf[m - 1], cf[m]);
                    }
                    //gfs.DrawLines(mypen, cf);
                    //gfs_e.DrawLines(mypen, f);
                    pos = i + 1;
                }
            }
            gfs.DrawImage(image, 0, 0, panel_Result.Width, panel_Result.Height);


            //int index = 0;
            //int code = m_BenchHandle.m_AccessHandle.GetIndexBaseKeyWord("TestResult_1", m_RecordId, ref index);
            //if (code != 1)
            //{
            //    MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);

            //object Obj = default(object);
            //code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "StartTime", ref Obj);
            //if (code != 1)
            //{
            //    return;
            //}
            startTime = m_HistoryStartTime;

            //code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "EndTime", ref Obj);
            //if (code != 1)
            //{
            //    return;
            //}
            endTime = m_HistoryEndTime;

            TimeSpan tp = endTime - startTime;
            int second = (int)tp.TotalSeconds;
            label_X1.Text = startTime.ToString("HH:mm:ss");
            textBox_StartTime.Text = label_X1.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7.Text = endTime.ToString("HH:mm:ss");
            textBox_EndTime.Text = label_X7.Text;

            UInt32 t1 = m_yMax * 1 / 5;
            label32.Text = t1.ToString() + "MPa";

            UInt32 t2 = m_yMax * 2 / 5;
            label33.Text = t2.ToString() + "MPa";

            UInt32 t3 = m_yMax * 3 / 5;
            label34.Text = t3.ToString() + "MPa";

            UInt32 t4 = m_yMax * 4 / 5;
            label36.Text = t4.ToString() + "MPa";
            //textBox_yMax.Text = m_yMax.ToString();
            //textBox_yMax_e.Text = m_yMax.ToString();
        }

        /// <summary>
        /// 将产品信息存入到本地缓存
        /// </summary>
        private string strSec = "TestInfoConfig"; //INI文件名
        private void ReadTextBoxData()
        {
            I_ProductName = TB_Name.Text;
            I_ProductNo = TB_ProductID.Text;
            I_ProductModel = TB_Module.Text;
            I_TestStandard = TB_Standard.Text;
            I_TestNiuju = TB_Niuju.Text;
            I_EquipmentNo = TB_EquipNum.Text;
            I_CompanyName = TB_CompanyName.Text;
            I_Operator = TB_Operator.Text;
            I_Tester = TB_Tester.Text;
            I_Confirmation = TB_Conformation.Text;

            WritePrivateProfileString(strSec, "ProductName", I_ProductName, strFilePath);
            WritePrivateProfileString(strSec, "ProductNo", I_ProductNo, strFilePath);
            WritePrivateProfileString(strSec, "ProductModel", I_ProductModel, strFilePath);
            WritePrivateProfileString(strSec, "TestStandard", I_TestStandard, strFilePath);
            WritePrivateProfileString(strSec, "TestNiuju", I_TestNiuju, strFilePath);
            WritePrivateProfileString(strSec, "EquipmentNo", I_EquipmentNo, strFilePath);
            WritePrivateProfileString(strSec, "CompanyName", I_CompanyName, strFilePath);
            WritePrivateProfileString(strSec, "Operator", I_Operator, strFilePath);
            WritePrivateProfileString(strSec, "Tester", I_Tester, strFilePath);
            WritePrivateProfileString(strSec, "Confirmation", I_Confirmation, strFilePath);
            WritePrivateProfileString(strSec, "Date", I_Date, strFilePath);
        }

        /// <summary>
        /// 将数据写入到数据库
        /// </summary>
        /// <returns></returns>
        private bool SaveDataToDataBase_M()
        {
            int code = 0;
            int index = 0;
            m_BenchHandle.m_AccessHandle.AddNewRow_M("TestResult_1", ref index);
            try
            {
                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "TestNo", m_TestNo);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                string StartTime = m_BenchHandle.m_StartTime.ToString("yyyy-MM-dd  HH:mm:ss");
                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "StartTime_s", StartTime);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "StartTime", m_BenchHandle.m_StartTime);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                DateTime endTime = default(DateTime);
                m_BenchHandle.ReadEndTime(ref endTime);
                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "EndTime", endTime);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "ProductName", I_ProductName);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "ProductNo", I_ProductNo);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "ProductModel", I_ProductModel);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "TestStandard", I_TestStandard);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "TestNiuju", I_TestNiuju);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "EquipmentNo", I_EquipmentNo);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "CompanyName", I_CompanyName);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "Operator", I_Operator);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "Tester", I_Tester);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "Confirmation", I_Confirmation);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "TestDate", I_Date);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                int rows = dataGridView_DateList.Rows.Count;
                float[] testDate = new float[(rows - 1) * 5];
                for (int i = 1; i < rows - 1; i++)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        if (j == 0)
                        {
                            string No = dataGridView_DateList[j, i].Value.ToString();
                            char[] a = No.ToCharArray();
                            string temp = "";
                            for (int m = 0; m < a.Length; m++)
                            {
                                if (a[m] == '-')
                                {
                                    continue;
                                }
                                temp += a[m];
                            }
                            testDate[(i - 1) * 5 + j] = Convert.ToSingle(temp);
                        }
                        else
                        {

                            testDate[(i - 1) * 5 + j] = Convert.ToSingle(dataGridView_DateList[j, i].Value.ToString());
                        }
                    }
                }

                //数量
                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "ResultsNum", rows - 2);//减去新建的行和标题行
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "xMax", (int)m_xMax);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData_M("TestResult_1", index, "yMax", m_yMax);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteFloatsData_M("TestResult_1", index, "TestResult", testDate);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }



                //读取曲线数据
                //int Num = 0;
                //int count = m_BenchHandle.m_PointFLists.Count;
                //for (int i = 0; i < count; i++)
                //{
                //    Num += m_BenchHandle.m_PointFLists[i].m_pts.Length;
                //    Num++;
                //}
                //float[] PointF_x = new float[Num];
                //float[] PointF_y = new float[Num];
                //int PointFcount = 0;    //float数组计数
                //for (int i = 0; i < count; i++)
                //{
                //    int len = m_BenchHandle.m_PointFLists[i].m_pts.Length;
                //    for (int j = 0; j < len; j++)
                //    {
                //        PointF_x[PointFcount + j] = m_BenchHandle.m_PointFLists[i].m_pts[j].X;
                //        PointF_y[PointFcount + j] = m_BenchHandle.m_PointFLists[i].m_pts[j].Y;
                //    }
                //    PointFcount += len;
                //    PointF_x[PointFcount] = 0;
                //    PointF_y[PointFcount] = 0;
                //    PointFcount++;
                //}

                int len = m_ChangedPointF.Length;
                float[] PointF_x = new float[len];
                float[] PointF_y = new float[len];
                for (int i = 0; i < len; i++)
                {
                    PointF_x[i] = m_ChangedPointF[i].X;
                    PointF_y[i] = m_ChangedPointF[i].Y;
                }

                code = m_BenchHandle.m_AccessHandle.WriteFloatsData_M("TestResult_1", index, "CuverData_X", PointF_x);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteFloatsData_M("TestResult_1", index, "CuverData_Y", PointF_y);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow_M("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.SaveDate_M();
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 将数据写入到数据库
        /// </summary>
        /// <returns></returns>
        private bool SaveDataToDataBase()
        {
            int code = 0;
            int index = 0;
            m_BenchHandle.m_AccessHandle.AddNewRow("TestResult_1", ref index);
            try
            {
                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "TestNo", m_TestNo);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                string StartTime = m_BenchHandle.m_StartTime.ToString("yyyy-MM-dd  HH:mm:ss");
                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "StartTime_s", StartTime);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "StartTime", m_BenchHandle.m_StartTime);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                DateTime endTime = default(DateTime);
                m_BenchHandle.ReadEndTime(ref endTime);
                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "EndTime", endTime);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "ProductName", I_ProductName);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "ProductNo", I_ProductNo);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "ProductModel", I_ProductModel);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "TestStandard", I_TestStandard);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "TestNiuju", I_TestNiuju);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "EquipmentNo", I_EquipmentNo);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "CompanyName", I_CompanyName);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "Operator", I_Operator);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "Tester", I_Tester);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "Confirmation", I_Confirmation);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "TestDate", I_Date);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                int rows = dataGridView_DateList.Rows.Count;
                float[] testDate = new float[(rows - 1) * 5];
                for (int i = 1; i < rows - 1; i++)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        if (j == 0)
                        {
                            string No = dataGridView_DateList[j, i].Value.ToString();
                            char[] a = No.ToCharArray();
                            string temp = "";
                            for (int m = 0; m < a.Length; m++)
                            {
                                if (a[m] == '-')
                                {
                                    continue;
                                }
                                temp += a[m];
                            }
                            testDate[(i - 1) * 5 + j] = Convert.ToSingle(temp);
                        }
                        else
                        {

                            testDate[(i - 1) * 5 + j] = Convert.ToSingle(dataGridView_DateList[j, i].Value.ToString());
                        }
                    }
                }

                //数量
                code = m_BenchHandle.m_AccessHandle.WriteSigleData("TestResult_1", index, "ResultsNum", rows - 2);//减去新建的行和标题行
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteFloatsData("TestResult_1", index, "TestResult", testDate);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                //读取曲线数据
                //int Num = 0;
                //int count = m_BenchHandle.m_PointFLists.Count;
                //for (int i = 0; i < count; i++)
                //{
                //    Num += m_BenchHandle.m_PointFLists[i].m_pts.Length;
                //    Num++;
                //}
                //float[] PointF_x = new float[Num];
                //float[] PointF_y = new float[Num];
                //int PointFcount = 0;    //float数组计数
                //for (int i = 0; i < count; i++)
                //{
                //    int len = m_BenchHandle.m_PointFLists[i].m_pts.Length;
                //    for (int j = 0; j < len; j++)
                //    {
                //        PointF_x[PointFcount + j] = m_BenchHandle.m_PointFLists[i].m_pts[j].X;
                //        PointF_y[PointFcount + j] = m_BenchHandle.m_PointFLists[i].m_pts[j].Y;
                //    }
                //    PointFcount += len;
                //    PointF_x[PointFcount] = 0;
                //    PointF_y[PointFcount] = 0;
                //    PointFcount++;
                //}

                int len = m_ChangedPointF.Length;
                float[] PointF_x = new float[len];
                float[] PointF_y = new float[len];
                for (int i = 0; i < len; i++)
                {
                    PointF_x[i] = m_ChangedPointF[i].X;
                    PointF_y[i] = m_ChangedPointF[i].Y;
                }

                code = m_BenchHandle.m_AccessHandle.WriteFloatsData("TestResult_1", index, "CuverData_X", PointF_x);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_BenchHandle.m_AccessHandle.WriteFloatsData("TestResult_1", index, "CuverData_Y", PointF_y);
                if (code != 1)
                {
                    m_BenchHandle.m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 创建数据报表
        /// </summary>
        /// <returns></returns>
        private string m_BaseFilePath = System.Windows.Forms.Application.StartupPath + "\\1#";
        private int D_delaytime = 1;
        private bool CreatOfficeReport()
        {
            string Name = DateTime.Now.ToString("yyyyMMddHHmmss") + ".doc";
            string FileName = m_BaseFilePath + "\\" + Name;
            //创建Word文档
            Object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document WordDoc = WordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            //设置格式
            WordApp.Selection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceAtLeast;//1倍行距
            WordApp.Selection.ParagraphFormat.SpaceBefore = float.Parse("0");
            WordApp.Selection.ParagraphFormat.SpaceBeforeAuto = 0;
            WordApp.Selection.ParagraphFormat.SpaceAfter = float.Parse("0");//段后间距
            WordApp.Selection.ParagraphFormat.SpaceAfterAuto = 0;
            WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            try
            {
                for (int i = 0; i < D_delaytime; i++) ;
                WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                string strContent = m_OfficeInfo + "\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;
                for (int i = 0; i < D_delaytime; i++) ;


                WordDoc.Paragraphs.Last.Range.Font.Name = "宋体";
                WordDoc.Paragraphs.Last.Range.Font.Size = 11;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "压力试验报告/Hydrostatic Test Chart\n";
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                //for (int i = 0; i < D_delaytime; i++) ;

                //WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                //WordDoc.Paragraphs.Last.Range.Font.Size = 10;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                //strContent = "Hydrostatic Test Chart\n";
                //WordDoc.Paragraphs.Last.Range.Text = strContent;

                for (int i = 0; i < D_delaytime; i++) ;

                WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;//水平居中
                WordDoc.Paragraphs.Last.Range.Font.Name = "宋体";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                strContent = "产品信息/Product Information\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                //for (int i = 0; i < D_delaytime; i++) ;

                //WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                //WordDoc.Paragraphs.Last.Range.Font.Size = 8;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                //strContent = "Product Information\n";
                //WordDoc.Paragraphs.Last.Range.Text = strContent;

                for (int i = 0; i < D_delaytime; i++) ;

                //添加表格
                Microsoft.Office.Interop.Word.Table table1 = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, 2, 7, ref oMissing, ref oMissing);
                //设置表格样式
                table1.Range.Font.Name = "宋体";
                table1.Range.Font.Size = 8;
                table1.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
                table1.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                table1.Cell(1, 1).Range.Text = "产品图号及名称/Draw No";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 2).Range.Text = "产品编号/Serial No";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 3).Range.Text = "试压标准/Test Standard";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 4).Range.Text = "扭矩/Test Torque(N.m)";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 5).Range.Text = "设备编号/Equipment No";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 6).Range.Text = "单位名称/Company Name";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 7).Range.Text = "合同号/Contract No";
                for (int i = 0; i < D_delaytime; i++) ;

                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 1).Range.Text = TB_Name.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 2).Range.Text = TB_ProductID.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 3).Range.Text = TB_Module.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 4).Range.Text = TB_Standard.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 5).Range.Text = TB_Niuju.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 6).Range.Text = TB_EquipNum.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 7).Range.Text = TB_CompanyName.Text;
                for (int i = 0; i < D_delaytime; i++) ;

                //插入文本，试验数据
                strContent = "\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;
                for (int i = 0; i < D_delaytime; i++) ;

                WordDoc.Paragraphs.Last.Range.Font.Name = "宋体";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "试验数据/Test Parameter\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;
                //for (int i = 0; i < D_delaytime; i++) ;

                //WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                //WordDoc.Paragraphs.Last.Range.Font.Size = 8;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                //strContent = "Test Parameter\n";
                //WordDoc.Paragraphs.Last.Range.Text = strContent;

                for (int i = 0; i < D_delaytime; i++) ;
                //第二个表格
                int rows = dataGridView_DateList.Rows.Count;
                Microsoft.Office.Interop.Word.Table table2 = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, rows - 1, 5, ref oMissing, ref oMissing);
                //表格样式
                table2.Range.Font.Size = 8;
                table2.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
                table2.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                //for (int i = 0; i < D_delaytime; i++) ;
                //table2.Cell(1, 1).Range.Text = "级数/段数";
                //for (int i = 0; i < D_delaytime; i++) ;
                //table2.Cell(1, 2).Range.Text = "初始压力（MPa）";
                //for (int i = 0; i < D_delaytime; i++) ;
                //table2.Cell(1, 3).Range.Text = "终止压力（MPa）";
                //for (int i = 0; i < D_delaytime; i++) ;
                //table2.Cell(1, 4).Range.Text = "保压时间（Min）";
                //for (int i = 0; i < D_delaytime; i++) ;
                //table2.Cell(1, 5).Range.Text = "保压压降（MPa）";
                //for (int i = 0; i < D_delaytime; i++) ;

                for (int i = 1; i < rows; i++)
                {
                    for (int j = 1; j <= 5; j++)
                    {
                        table2.Cell(i, j).Range.Text = dataGridView_DateList[j - 1, i - 1].Value.ToString();
                        //for (int t = 0; t < D_delaytime; t++) ;
                    }
                }

                for (int i = 0; i < D_delaytime; i++) ;

                strContent = "\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                //插入图片
                FileName = m_BaseFilePath + "\\Presure.png";//图片所在路径
                object LinkToFile = false;
                object SaveWithDocument = true;
                object Anchor = WordDoc.Paragraphs.Last.Range;
                //object Anchor = WordDoc.Application.Selection.Range;
                //WordDoc.Paragraphs.Last.Range = WordDoc.Application.Selection.Range;

                WordDoc.Application.ActiveDocument.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Anchor);
                Microsoft.Office.Interop.Word.Shape s = WordDoc.Application.ActiveDocument.InlineShapes[1].ConvertToShape();
                s.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapInline;
                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;

                object oEndOfDoc = "\\endofdoc";
                Microsoft.Office.Interop.Word.Paragraph return_pragraph;
                object myrange2 = WordDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                return_pragraph = WordDoc.Content.Paragraphs.Add(ref myrange2);
                return_pragraph.Range.InsertParagraphAfter(); //插入一个空白行  

                for (int i = 0; i < D_delaytime; i++) ;
                //插入试验结果
                WordDoc.Paragraphs.Last.Range.Font.Name = "宋体";
                WordDoc.Paragraphs.Last.Range.Font.Size = 11;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "试验结果：在保压周期内产品无可见渗漏\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                for (int i = 0; i < D_delaytime; i++) ;
                //插入试验结果
                WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "Test Result：no visibale leakage during each holding period\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                for (int i = 0; i < D_delaytime; i++) ;
                //第三个表格
                Microsoft.Office.Interop.Word.Table table3 = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, 2, 4, ref oMissing, ref oMissing);
                //表格样式
                table3.Range.Font.Name = "宋体";
                table3.Range.Font.Size = 8;
                table3.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
                table3.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(1, 1).Range.Text = "试验人员/Tested by";
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(1, 2).Range.Text = "检验人员/Inspected by";
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(1, 3).Range.Text = "审核/Review by";
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(1, 4).Range.Text = "日期/Date";
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(2, 1).Range.Text = TB_Operator.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(2, 2).Range.Text = TB_Tester.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(2, 3).Range.Text = TB_Conformation.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(2, 4).Range.Text = TB_Date.Text;
                for (int i = 0; i < D_delaytime; i++) ;


                //format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;
                object filename = m_BaseFilePath + "\\Report\\" + Name;
                for (int i = 0; i < D_delaytime; i++) ;

                WordDoc.SaveAs(ref filename, ref format, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                //关闭wordDoc 文档对象
                WordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                //关闭wordApp 组件对象
                WordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                FileName = filename.ToString();
            }
            catch (Exception)
            {
                //关闭wordDoc 文档对象
                WordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                //关闭wordApp 组件对象
                WordApp.Quit(ref oMissing, ref oMissing, ref oMissing);

                return false;
            }

            if (File.Exists(FileName))
            {
                //File.Open(FileName,FileMode.Open);
                System.Diagnostics.Process.Start(FileName);
            }


            return true;
        }

        private PointF[] m_HistoryPointF;
        private DateTime m_HistoryStartTime;
        private DateTime m_HistoryEndTime;
        private bool ReadCurveDataFromDataBase()
        {
            int code = 0;
            int index = 0;
            code = m_BenchHandle.m_AccessHandle.GetIndexBaseKeyWord("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            float[] X;
            float[] Y;

            code = m_BenchHandle.m_AccessHandle.ReadFloatsData("TestResult_1", index, "CuverData_X", out X);
            if (code != 1)
            {
                return false;
            }

            code = m_BenchHandle.m_AccessHandle.ReadFloatsData("TestResult_1", index, "CuverData_Y", out Y);
            if (code != 1)
            {
                return false;
            }

            PointF[] p = new PointF[X.Length];
            m_HistoryPointF = new PointF[X.Length];
            int len = X.Length;
            for (int i = 0; i < len; i++)
            {
                p[i].X = X[i];
                p[i].Y = Y[i];
                m_HistoryPointF[i].X = X[i];
                m_HistoryPointF[i].Y = Y[i];
            }

            //绘制曲线
            //m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result);
            //Graphics gfs = panel_Result.CreateGraphics();
            //Pen mypen = new Pen(Color.Red, 1.5f);

            //int pos = 0;
            //for (int i = 0; i < len; i++)
            //{
            //    if (i != 0 && p[i].X == 0 && p[i].Y == 0)
            //    {
            //        PointF[] f = new PointF[i - pos];
            //        for (int j = pos; j < i; j++)
            //        {
            //            f[j - pos] = p[j];
            //        }
            //        gfs.DrawLines(mypen, f);
            //        pos = i + 1;
            //    }
            //}

            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);

            object Obj = default(object);
            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "StartTime", ref Obj);
            if (code != 1)
            {
                return false;
            }
            startTime = Convert.ToDateTime(Obj);
            m_HistoryStartTime = startTime;

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "EndTime", ref Obj);
            if (code != 1)
            {
                return false;
            }
            endTime = Convert.ToDateTime(Obj);
            m_HistoryEndTime = endTime;

            label_X1.Text = startTime.ToString("HH:mm:ss");
            textBox_StartTime.Text = label_X1.Text;
            label_X1_e.Text = startTime.ToString("HH:mm:ss");
            textBox_StartTime_e.Text = label_X1_e.Text;
            //label_X7.Text = endTime.ToString("HH:mm:ss");
            //label_X7_e.Text = endTime.ToString("HH:mm:ss");

            //TimeSpan tp = endTime - startTime;
            //int second = (int)tp.TotalSeconds;
            //label_X1.Text = startTime.ToString("HH:mm:ss");

            //int H = 0;
            //int M = 0;
            //int S = 0;
            //GetlableTime(second / 6, ref H, ref M, ref S);
            //label_X2.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            //GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            //label_X3.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            //GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            //label_X4.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            //GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            //label_X5.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            //GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            //label_X6.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            return true;
        }

        private bool ReadCurveDataFromDataBase_M()
        {
            int code = 0;
            int index = 0;
            code = m_BenchHandle.m_AccessHandle.GetIndexBaseKeyWord_M("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            float[] X;
            float[] Y;

            code = m_BenchHandle.m_AccessHandle.ReadFloatsData_M("TestResult_1", index, "CuverData_X", out X);
            if (code != 1)
            {
                return false;
            }

            code = m_BenchHandle.m_AccessHandle.ReadFloatsData_M("TestResult_1", index, "CuverData_Y", out Y);
            if (code != 1)
            {
                return false;
            }

            PointF[] p = new PointF[X.Length];
            m_HistoryPointF = new PointF[X.Length];
            int len = X.Length;
            for (int i = 0; i < len; i++)
            {
                p[i].X = X[i];
                p[i].Y = Y[i];
                m_HistoryPointF[i].X = X[i];
                m_HistoryPointF[i].Y = Y[i];
            }

            //绘制曲线
            //m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result);
            //Graphics gfs = panel_Result.CreateGraphics();
            //Pen mypen = new Pen(Color.Red, 1.5f);

            //int pos = 0;
            //for (int i = 0; i < len; i++)
            //{
            //    if (i != 0 && p[i].X == 0 && p[i].Y == 0)
            //    {
            //        PointF[] f = new PointF[i - pos];
            //        for (int j = pos; j < i; j++)
            //        {
            //            f[j - pos] = p[j];
            //        }
            //        gfs.DrawLines(mypen, f);
            //        pos = i + 1;
            //    }
            //}

            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);

            object Obj = default(object);
            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "StartTime", ref Obj);
            if (code != 1)
            {
                return false;
            }
            startTime = Convert.ToDateTime(Obj);
            m_HistoryStartTime = startTime;

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "EndTime", ref Obj);
            if (code != 1)
            {
                return false;
            }
            endTime = Convert.ToDateTime(Obj);
            m_HistoryEndTime = endTime;

            //label_X1.Text = startTime.ToString("HH:mm:ss");
            //textBox_StartTime.Text = label_X1.Text;
            //label_X1_e.Text = startTime.ToString("HH:mm:ss");
            //textBox_StartTime_e.Text = label_X1_e.Text;
            //label_X7.Text = endTime.ToString("HH:mm:ss");
            //label_X7_e.Text = endTime.ToString("HH:mm:ss");

            //TimeSpan tp = endTime - startTime;
            //int second = (int)tp.TotalSeconds;
            //label_X1.Text = startTime.ToString("HH:mm:ss");

            //int H = 0;
            //int M = 0;
            //int S = 0;
            //GetlableTime(second / 6, ref H, ref M, ref S);
            //label_X2.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            //GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            //label_X3.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            //GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            //label_X4.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            //GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            //label_X5.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            //GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            //label_X6.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            return true;
        }

        /// <summary>
        /// 退出报表界面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Return_Click(object sender, EventArgs e)
        {
            //m_DrawCurve.Stop();
            this.Close();
            //this.Dispose();
        }

        /// <summary>
        /// 生成报表，中文
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private bool m_isCreatReport = false;
        private void button_CreatReport_Click(object sender, EventArgs e)
        {
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            bool ret = false;

            //ret = SavePic();
            //if (!ret)
            //{
            //    MessageBox.Show("创建报表失败2,请检查数据信息", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}
            SavePic();
            ReadTextBoxData();

            if (!m_isOpenDataBase)
            {
                if (!m_isCreatReport)
                {
                    ret = SaveDataToDataBase_M();
                    if (!ret)
                    {
                        MessageBox.Show("创建报表失败,请检查数据信息", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Cursor = System.Windows.Forms.Cursors.Arrow;
                        return;
                    }
                    m_isCreatReport = true;
                }
            }

            ret = CreatOfficeReport();
            if (!ret)
            {
                MessageBox.Show("创建报表失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            //m_DrawCurve.Start();
        }

        //读取历史数据
        private SelectHistoryData_No1 m_SelectHistoryData_No1Handle;
        public bool m_isNoSelect = true;
        public int m_RecordId = 0;
        private bool m_isOpenDataBase = false;
        private void button_Save_Path_Click(object sender, EventArgs e)
        {
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            m_SelectHistoryData_No1Handle = new SelectHistoryData_No1(this, m_BenchHandle);
            m_SelectHistoryData_No1Handle.ControlBox = false;
            m_SelectHistoryData_No1Handle.ShowDialog();

            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            if (m_isNoSelect)
            {
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            bool ret = ReadInfoDataFromDataBase_M();
            if (!ret)
            {
                MessageBox.Show("读取设备信息失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            ret = ReadTestDataFromDataBase_M();
            if (!ret)
            {
                MessageBox.Show("读取试验数据失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            ret = ReadCurveDataFromDataBase_M();
            if (!ret)
            {
                MessageBox.Show("读取曲线信息失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            //m_DrawCurve.Stop();
            m_isOpenDataBase = true;
            //button_CreatReport.Enabled = false;
            //button2.Enabled = false;
            //m_DrawHistoryCurveTimer.Start();
            DrawHistoryCurve();
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
        }

        #endregion

        #region 公共区

        //读取试验信息
        private string strFilePath = System.Windows.Forms.Application.StartupPath + @"\1#\TestInfoConfig.ini";//获取INI文件路径

        /// <summary>
        /// 读取测试信息，如测试人员--公共
        /// </summary>        
        private void ReadTestInfo()
        {
            string strSecfileName = Path.GetFileNameWithoutExtension(strFilePath);
            I_ProductName = ContentValue(strSecfileName, "ProductName");
            TB_Name.Text = I_ProductName;
            TB_Name_e.Text = I_ProductName;

            I_ProductNo = ContentValue(strSecfileName, "ProductNo");
            TB_ProductID.Text = I_ProductNo;
            TB_ProductID_e.Text = I_ProductNo;

            I_ProductModel = ContentValue(strSecfileName, "ProductModel");
            TB_Module.Text = I_ProductModel;
            TB_Module_e.Text = I_ProductModel;

            I_TestStandard = ContentValue(strSecfileName, "TestStandard");
            TB_Standard.Text = I_TestStandard;
            TB_Standard_e.Text = I_TestStandard;

            I_TestNiuju = ContentValue(strSecfileName, "TestNiuju");
            TB_Niuju.Text = I_TestNiuju;
            TB_Niuju_e.Text = I_TestNiuju;

            I_EquipmentNo = ContentValue(strSecfileName, "EquipmentNo");
            TB_EquipNum.Text = I_EquipmentNo;
            TB_EquipNum_e.Text = I_EquipmentNo;

            I_CompanyName = ContentValue(strSecfileName, "CompanyName");
            TB_CompanyName.Text = I_CompanyName;
            TB_CompanyName_e.Text = I_CompanyName;

            I_Operator = ContentValue(strSecfileName, "Operator");
            TB_Operator.Text = I_Operator;
            TB_Operator_e.Text = I_Operator;

            I_Tester = ContentValue(strSecfileName, "Tester");
            TB_Tester.Text = I_Tester;
            TB_Tester_e.Text = I_Tester;

            I_Confirmation = ContentValue(strSecfileName, "Confirmation");
            TB_Conformation.Text = I_Confirmation;
            TB_Conformation_e.Text = I_Confirmation;

            I_Date = ContentValue(strSecfileName, "Date");
            TB_Date.Text = I_Date;
            TB_Date_e.Text = I_Date;
        }
        private string ContentValue(string Section, string key)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }

        /// <summary>
        /// 读取试验数据
        /// </summary>
        private void ReadTestDate()
        {
            dataGridView_DateList.ColumnCount = 5;
            dataGridView_DateList_e.ColumnCount = 5;
            int len = m_BenchHandle.m_TestResultLists.Count;
            for (int i = 0; i < len + 1; i++)    //
            {
                dataGridView_DateList.Rows.Add();
                dataGridView_DateList_e.Rows.Add();
            }
            dataGridView_DateList.Rows[0].Cells[0].Value = "级数/段数/Stage No";
            dataGridView_DateList.Rows[0].Cells[1].Value = "初始压力/Start Pressure（MPa）";
            dataGridView_DateList.Rows[0].Cells[2].Value = "终止压力/Final Pressure（MPa）";
            dataGridView_DateList.Rows[0].Cells[3].Value = "保压时间/Hold Period（Min）";
            dataGridView_DateList.Rows[0].Cells[4].Value = "保压压降/Pressure Reduction（MPa）";

            dataGridView_DateList_e.Rows[0].Cells[0].Value = "Stage No";
            dataGridView_DateList_e.Rows[0].Cells[1].Value = "Start Pressure(MPa)";
            dataGridView_DateList_e.Rows[0].Cells[2].Value = "Final Pressure(MPa)";
            dataGridView_DateList_e.Rows[0].Cells[3].Value = "Hold Period(Min)";
            dataGridView_DateList_e.Rows[0].Cells[4].Value = "Pressure Reduction(MPa)";

            for (int i = 1; i < len + 1; i++)
            {
                //中文报告
                dataGridView_DateList.Rows[i].Cells[0].Value = m_BenchHandle.m_TestResultLists[i - 1].m_No; ;
                dataGridView_DateList.Rows[i].Cells[1].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_InitPressure.ToString("0.00"));
                dataGridView_DateList.Rows[i].Cells[2].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_EndPressure.ToString("0.00"));
                dataGridView_DateList.Rows[i].Cells[3].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_KeepTime.ToString("0.00"));
                dataGridView_DateList.Rows[i].Cells[4].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_DropPressure.ToString("0.00"));
                //英文报告
                dataGridView_DateList_e.Rows[i].Cells[0].Value = m_BenchHandle.m_TestResultLists[i - 1].m_No;
                dataGridView_DateList_e.Rows[i].Cells[1].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_InitPressure.ToString("0.00"));
                dataGridView_DateList_e.Rows[i].Cells[2].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_EndPressure.ToString("0.00"));
                dataGridView_DateList_e.Rows[i].Cells[3].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_KeepTime.ToString("0.00"));
                dataGridView_DateList_e.Rows[i].Cells[4].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_DropPressure.ToString("0.00"));
            }
        }

        //坐标转换，根据固定最大值（xMax，yMax）来绘制图，
        private void PointTranslate(PointF[] p, out PointF[] f, Panel panel, float xMax, float yMax)
        {
            int Length = 0;
            int MaxSize_x = panel.Size.Width;
            int MaxSize_y = panel.Size.Height - 10;
            Length = p.Length;
            f = new PointF[Length];
            for (int i = 0; i < Length; i++)
            {
                f[i].X = (p[i].X / xMax) * MaxSize_x;
                f[i].Y = ((yMax - p[i].Y) / yMax) * MaxSize_y + 10;
            }
        }
        /// <summary>
        /// 从数据库中读取产品信息
        /// </summary>
        /// <returns></returns>
        private bool ReadInfoDataFromDataBase()
        {
            int code = 0;
            int index = 0;
            code = m_BenchHandle.m_AccessHandle.GetIndexBaseKeyWord("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            object Obj = default(object);
            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "ProductName", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_ProductName = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "ProductNo", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_ProductNo = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "ProductModel", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_ProductModel = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "TestStandard", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_TestStandard = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "TestNiuju", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_TestNiuju = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "EquipmentNo", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_EquipmentNo = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "CompanyName", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_CompanyName = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "Operator", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Operator = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "Tester", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Tester = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "Confirmation", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Confirmation = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "TestDate", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Date = Obj.ToString();

            TB_Name.Text = I_ProductName;
            TB_ProductID.Text = I_ProductNo;
            TB_Module.Text = I_ProductModel;
            TB_Standard.Text = I_TestStandard;
            TB_Niuju.Text = I_TestNiuju;
            TB_EquipNum.Text = I_EquipmentNo;
            TB_CompanyName.Text = I_CompanyName;
            TB_Operator.Text = I_Operator;
            TB_Tester.Text = I_Tester;
            TB_Conformation.Text = I_Confirmation;

            TB_Name_e.Text = I_ProductName;
            TB_ProductID_e.Text = I_ProductNo;
            TB_Module_e.Text = I_ProductModel;
            TB_Standard_e.Text = I_TestStandard;
            TB_Niuju_e.Text = I_TestNiuju;
            TB_EquipNum_e.Text = I_EquipmentNo;
            TB_CompanyName_e.Text = I_CompanyName;
            TB_Operator_e.Text = I_Operator;
            TB_Tester_e.Text = I_Tester;
            TB_Conformation_e.Text = I_Confirmation;

            TB_Result.Text = "在保压周期内无可见渗漏";
            TB_Result_e.Text = "no visibale leakage during each holding period";
            return true;
        }

        private bool ReadInfoDataFromDataBase_M()
        {
            int code = 0;
            int index = 0;
            code = m_BenchHandle.m_AccessHandle.GetIndexBaseKeyWord_M("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            object Obj = default(object);
            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "ProductName", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_ProductName = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "ProductNo", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_ProductNo = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "ProductModel", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_ProductModel = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "TestStandard", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_TestStandard = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "TestNiuju", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_TestNiuju = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "EquipmentNo", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_EquipmentNo = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "CompanyName", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_CompanyName = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "Operator", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Operator = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "Tester", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Tester = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "Confirmation", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Confirmation = Obj.ToString();

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "TestDate", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Date = Obj.ToString();

            TB_Name.Text = I_ProductName;
            TB_ProductID.Text = I_ProductNo;
            TB_Module.Text = I_ProductModel;
            TB_Standard.Text = I_TestStandard;
            TB_Niuju.Text = I_TestNiuju;
            TB_EquipNum.Text = I_EquipmentNo;
            TB_CompanyName.Text = I_CompanyName;
            TB_Operator.Text = I_Operator;
            TB_Tester.Text = I_Tester;
            TB_Conformation.Text = I_Confirmation;

            TB_Name_e.Text = I_ProductName;
            TB_ProductID_e.Text = I_ProductNo;
            TB_Module_e.Text = I_ProductModel;
            TB_Standard_e.Text = I_TestStandard;
            TB_Niuju_e.Text = I_TestNiuju;
            TB_EquipNum_e.Text = I_EquipmentNo;
            TB_CompanyName_e.Text = I_CompanyName;
            TB_Operator_e.Text = I_Operator;
            TB_Tester_e.Text = I_Tester;
            TB_Conformation_e.Text = I_Confirmation;

            TB_Result.Text = "在保压周期内无可见渗漏";
            TB_Result_e.Text = "no visibale leakage during each holding period";
            return true;
        }

        /// <summary>
        /// 从数据库中读取测试数据
        /// </summary>
        /// <returns></returns>
        private bool ReadTestDataFromDataBase()
        {
            int code = 0;
            int index = 0;
            code = m_BenchHandle.m_AccessHandle.GetIndexBaseKeyWord("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            int TestDataNum = 0;
            object Obj = default(object);
            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "ResultsNum", ref Obj);
            if (code != 1)
            {
                return false;
            }
            TestDataNum = Convert.ToInt32(Obj);

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "xMax", ref Obj);
            if (code != 1)
            {
                return false;
            }
            m_xMax = Convert.ToInt32(Obj);

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "yMax", ref Obj);
            if (code != 1)
            {
                return false;
            }
            m_yMax = Convert.ToUInt32(Obj);
            textBox_yMax.Text = m_yMax.ToString();
            textBox_yMax_e.Text = m_yMax.ToString();

            float[] TestData;
            code = m_BenchHandle.m_AccessHandle.ReadFloatsData("TestResult_1", index, "TestResult", out TestData);
            if (code != 1)
            {
                return false;
            }

            dataGridView_DateList.Rows.Clear();
            dataGridView_DateList_e.Rows.Clear();
            dataGridView_DateList.ColumnCount = 5;
            dataGridView_DateList_e.ColumnCount = 5;
            for (int i = 0; i < TestDataNum + 1; i++)    //
            {
                dataGridView_DateList.Rows.Add();
                dataGridView_DateList_e.Rows.Add();
            }
            dataGridView_DateList.Rows[0].Cells[0].Value = "级数/段数/Stage No";
            dataGridView_DateList.Rows[0].Cells[1].Value = "初始压力/Start Pressure（MPa）";
            dataGridView_DateList.Rows[0].Cells[2].Value = "终止压力/Final Pressure（MPa）";
            dataGridView_DateList.Rows[0].Cells[3].Value = "保压时间/Hold Period（Min）";
            dataGridView_DateList.Rows[0].Cells[4].Value = "保压压降/Pressure Reduction（MPa）";

            dataGridView_DateList_e.Rows[0].Cells[0].Value = "Stage No";
            dataGridView_DateList_e.Rows[0].Cells[1].Value = "Start Pressure(MPa)";
            dataGridView_DateList_e.Rows[0].Cells[2].Value = "Final Pressure(MPa)";
            dataGridView_DateList_e.Rows[0].Cells[3].Value = "Hold Period(Min)";
            dataGridView_DateList_e.Rows[0].Cells[4].Value = "Pressure Reduction(MPa)";

            for (int i = 1; i < TestDataNum + 1; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    if (j == 0)
                    {
                        string No = ((int)TestData[(i - 1) * 5 + j]).ToString();
                        char[] a = No.ToCharArray();
                        if (a.Length == 2)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + "--" + a[1];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + "--" + a[1];
                        }
                        if (a.Length == 3)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + a[1] + "--" + a[2];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + a[1] + "--" + a[2];

                        }
                        if (a.Length == 4)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + "--" + a[3];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + "--" + a[3];
                        }
                        if (a.Length == 5)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + a[3] + "--" + a[4];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + a[3] + "--" + a[4];
                        }
                    }
                    else
                    {
                        dataGridView_DateList.Rows[i].Cells[j].Value = TestData[(i - 1) * 5 + j];
                        dataGridView_DateList_e.Rows[i].Cells[j].Value = TestData[(i - 1) * 5 + j];
                    }
                }
            }

            return true;
        }

        private bool ReadTestDataFromDataBase_M()
        {
            int code = 0;
            int index = 0;
            code = m_BenchHandle.m_AccessHandle.GetIndexBaseKeyWord_M("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            int TestDataNum = 0;
            object Obj = default(object);
            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "ResultsNum", ref Obj);
            if (code != 1)
            {
                return false;
            }
            TestDataNum = Convert.ToInt32(Obj);

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "xMax", ref Obj);
            if (code != 1)
            {
                return false;
            }
            m_xMax = Convert.ToInt32(Obj);

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "yMax", ref Obj);
            if (code != 1)
            {
                return false;
            }
            m_yMax = Convert.ToUInt32(Obj);
            textBox_yMax.Text = m_yMax.ToString();
            textBox_yMax_e.Text = m_yMax.ToString();

            float[] TestData;
            code = m_BenchHandle.m_AccessHandle.ReadFloatsData_M("TestResult_1", index, "TestResult", out TestData);
            if (code != 1)
            {
                return false;
            }

            dataGridView_DateList.Rows.Clear();
            dataGridView_DateList_e.Rows.Clear();
            dataGridView_DateList.ColumnCount = 5;
            dataGridView_DateList_e.ColumnCount = 5;
            for (int i = 0; i < TestDataNum + 1; i++)    //
            {
                dataGridView_DateList.Rows.Add();
                dataGridView_DateList_e.Rows.Add();
            }
            dataGridView_DateList.Rows[0].Cells[0].Value = "级数/段数/Stage No";
            dataGridView_DateList.Rows[0].Cells[1].Value = "初始压力/Start Pressure（MPa）";
            dataGridView_DateList.Rows[0].Cells[2].Value = "终止压力/Final Pressure（MPa）";
            dataGridView_DateList.Rows[0].Cells[3].Value = "保压时间/Hold Period（Min）";
            dataGridView_DateList.Rows[0].Cells[4].Value = "保压压降/Pressure Reduction（MPa）";

            dataGridView_DateList_e.Rows[0].Cells[0].Value = "Stage No";
            dataGridView_DateList_e.Rows[0].Cells[1].Value = "Start Pressure(MPa)";
            dataGridView_DateList_e.Rows[0].Cells[2].Value = "Final Pressure(MPa)";
            dataGridView_DateList_e.Rows[0].Cells[3].Value = "Hold Period(Min)";
            dataGridView_DateList_e.Rows[0].Cells[4].Value = "Pressure Reduction(MPa)";

            for (int i = 1; i < TestDataNum + 1; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    if (j == 0)
                    {
                        string No = ((int)TestData[(i - 1) * 5 + j]).ToString();
                        char[] a = No.ToCharArray();
                        if (a.Length == 2)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + "--" + a[1];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + "--" + a[1];
                        }
                        if (a.Length == 3)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + a[1] + "--" + a[2];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + a[1] + "--" + a[2];

                        }
                        if (a.Length == 4)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + "--" + a[3];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + "--" + a[3];
                        }
                        if (a.Length == 5)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + a[3] + "--" + a[4];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + a[3] + "--" + a[4];
                        }
                    }
                    else
                    {
                        dataGridView_DateList.Rows[i].Cells[j].Value = TestData[(i - 1) * 5 + j];
                        dataGridView_DateList_e.Rows[i].Cells[j].Value = TestData[(i - 1) * 5 + j];
                    }
                }
            }

            return true;
        }

        private void TB_Name_TextChanged(object sender, EventArgs e)
        {
            TB_Name_e.Text = TB_Name.Text;
        }

        private void TB_Name_e_TextChanged(object sender, EventArgs e)
        {
            TB_Name.Text = TB_Name_e.Text;
        }

        private void TB_ProductID_TextChanged(object sender, EventArgs e)
        {
            TB_ProductID_e.Text = TB_ProductID.Text;
        }

        private void TB_ProductID_e_TextChanged(object sender, EventArgs e)
        {
            TB_ProductID.Text = TB_ProductID_e.Text;
        }

        private void TB_Module_TextChanged(object sender, EventArgs e)
        {
            TB_Module_e.Text = TB_Module.Text;
        }

        private void TB_Module_e_TextChanged(object sender, EventArgs e)
        {
            TB_Module.Text = TB_Module_e.Text;
        }

        private void TB_Standard_TextChanged(object sender, EventArgs e)
        {
            TB_Standard_e.Text = TB_Standard.Text;
        }

        private void TB_Standard_e_TextChanged(object sender, EventArgs e)
        {
            TB_Standard.Text = TB_Standard_e.Text;
        }

        private void TB_Niuju_TextChanged(object sender, EventArgs e)
        {
            TB_Niuju_e.Text = TB_Niuju.Text;
        }

        private void TB_Niuju_e_TextChanged(object sender, EventArgs e)
        {
            TB_Niuju.Text = TB_Niuju_e.Text;
        }

        private void TB_EquipNum_TextChanged(object sender, EventArgs e)
        {
            TB_EquipNum_e.Text = TB_EquipNum.Text;
        }

        private void TB_EquipNum_e_TextChanged(object sender, EventArgs e)
        {
            TB_EquipNum.Text = TB_EquipNum_e.Text;
        }

        private void TB_CompanyName_TextChanged(object sender, EventArgs e)
        {
            TB_CompanyName_e.Text = TB_CompanyName.Text;
        }

        private void TB_CompanyName_e_TextChanged(object sender, EventArgs e)
        {
            TB_CompanyName.Text = TB_CompanyName_e.Text;
        }

        private void TB_Operator_TextChanged(object sender, EventArgs e)
        {
            TB_Operator_e.Text = TB_Operator.Text;
        }

        private void TB_Operator_e_TextChanged(object sender, EventArgs e)
        {
            TB_Operator.Text = TB_Operator_e.Text;
        }

        private void TB_Tester_TextChanged(object sender, EventArgs e)
        {
            TB_Tester_e.Text = TB_Tester.Text;
        }

        private void TB_Tester_e_TextChanged(object sender, EventArgs e)
        {
            TB_Tester.Text = TB_Tester_e.Text;
        }

        private void TB_Conformation_TextChanged(object sender, EventArgs e)
        {
            TB_Conformation_e.Text = TB_Conformation.Text;
        }

        private void TB_Conformation_e_TextChanged(object sender, EventArgs e)
        {
            TB_Conformation.Text = TB_Conformation_e.Text;
        }
        #endregion

        #region 英文区

        private void DrawCurveFun_e(object o, ElapsedEventArgs e)
        {
            ReadTestPic1_e();
        }

        /// <summary>
        /// 将内存数据绘制成报表中的数据
        /// </summary>
        private void ReadTestPic1_e()
        {
            //绘制底层模板
            //m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result);
            m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result_e);

            //
            Graphics gfs_e = panel_Result_e.CreateGraphics();
            Pen mypen_e = new Pen(Color.Red, 1.5f);
            Image image = new Bitmap(panel_Result_e.Width, panel_Result_e.Height);
            Graphics g = Graphics.FromImage(image);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            //去横轴最大值
            int count = m_BenchHandle.m_PointFLists.Count;
            if (count <= 0)
            {
                return;
            }
            float xMax = 0;

            int PointCount = 0;
            for (int i = 0; i < count; i++)
            {
                int len = m_BenchHandle.m_PointFLists[i].m_pts.Length;
                //if (len <= 0)
                //{
                //    return;
                //}
                //xMax += m_BenchHandle.m_PointFLists[i].m_pts[len - 1].X + 10;
                PointCount += len;
            }
            xMax += 2;

            int length = m_BenchHandle.m_PointFLists[count - 1].m_pts.Length;
            if (length <= 0)
            {
                return;
            }
            xMax = m_BenchHandle.m_PointFLists[count - 1].m_pts[0].X + m_BenchHandle.m_PointFLists[count - 1].m_pts[length - 1].X;
            xMax += 2;

            m_xMax = xMax;

            //float position = 0;

            m_ChangedPointF = new PointF[PointCount + count];
            int ChangedPontfCount = 0;
            //坐标变换
            for (int i = 0; i < count; i++)
            {
                int len = m_BenchHandle.m_PointFLists[i].m_pts.Length;
                if (len <= 0)
                {
                    continue;
                }
                PointF[] temp = new PointF[len];
                PointF[] Ctemp;

                for (int j = 0; j < len; j++)
                {
                    if (j == 0)
                    {
                        temp[j].X = m_BenchHandle.m_PointFLists[i].m_pts[0].X;
                        temp[j].Y = m_BenchHandle.m_PointFLists[i].m_pts[j].Y;
                    }
                    else
                    {
                        temp[j].X = m_BenchHandle.m_PointFLists[i].m_pts[0].X + m_BenchHandle.m_PointFLists[i].m_pts[j].X;
                        temp[j].Y = m_BenchHandle.m_PointFLists[i].m_pts[j].Y;
                    }
                }

                PointTranslate(temp, out Ctemp, panel_Result_e, xMax, m_yMax);
                for (int m = 1; m < Ctemp.Length; m++)
                {
                    g.DrawLine(mypen_e, Ctemp[m - 1], Ctemp[m]);
                }

                //gfs_e.DrawLines(mypen_e, Ctemp);

                //position = temp[len - 1].X;

                for (int m = 0; m < Ctemp.Length; m++)
                {
                    m_ChangedPointF[ChangedPontfCount + m] = temp[m];
                }
                ChangedPontfCount += len;
                m_ChangedPointF[ChangedPontfCount].X = 0;
                m_ChangedPointF[ChangedPontfCount].Y = 0;
                ChangedPontfCount++;
            }
            gfs_e.DrawImage(image, 0, 0, panel_Result_e.Width, panel_Result_e.Height);

            DateTime StartTime = m_BenchHandle.m_StartTime;
            DateTime EndTime = default(DateTime);
            m_BenchHandle.ReadEndTime(ref EndTime);
            TimeSpan tp = EndTime - StartTime;
            int second = (int)tp.TotalSeconds;
            label_X1_e.Text = StartTime.ToString("HH:mm:ss");
            textBox_StartTime_e.Text = label_X1_e.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7_e.Text = EndTime.ToString("HH:mm:ss");
            textBox_EndTime_e.Text = label_X7_e.Text;

            UInt32 t1 = m_yMax * 1 / 5;
            label45.Text = t1.ToString() + "MPa";

            UInt32 t2 = m_yMax * 2 / 5;
            label44.Text = t2.ToString() + "MPa";

            UInt32 t3 = m_yMax * 3 / 5;
            label43.Text = t3.ToString() + "MPa";

            UInt32 t4 = m_yMax * 4 / 5;
            label41.Text = t4.ToString() + "MPa";
            //textBox_yMax.Text = m_yMax.ToString();
            //textBox_yMax_e.Text = m_yMax.ToString();
        }

        /// <summary>
        /// 英文报表曲线绘制定时函数
        /// </summary>
        /// <param name="o"></param>
        /// <param name="e"></param>
        private void DrawHistoryCurveFun_e(object o, ElapsedEventArgs e)
        {
            DrawHistoryCurve_e();
            SavePic();
            //ReadTextBoxData_e();
        }

        /// <summary>
        /// 绘制历史曲线--英文报告
        /// </summary>
        private void DrawHistoryCurve_e()
        {
            m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result_e);

            Graphics gfs_e = panel_Result_e.CreateGraphics();
            Pen mypen_e = new Pen(Color.Red, 1.5f);
            Image image = new Bitmap(panel_Result_e.Width, panel_Result_e.Height);
            Graphics g = Graphics.FromImage(image);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            int len = m_HistoryPointF.Length;

            int pos = 0;
            for (int i = 0; i < len; i++)
            {
                if (i != 0 && m_HistoryPointF[i].X == 0 && m_HistoryPointF[i].Y == 0)
                {
                    PointF[] f = new PointF[i - pos];
                    for (int j = pos; j < i; j++)
                    {
                        f[j - pos].X = m_HistoryPointF[j].X;
                        f[j - pos].Y = m_HistoryPointF[j].Y;
                    }
                    PointF[] cf;
                    PointTranslate(f, out cf, panel_Result_e, m_xMax, m_yMax);
                    for (int m = 1; m < cf.Length; m++)
                    {
                        gfs_e.DrawLine(mypen_e, cf[m - 1], cf[m]);
                    }
                    g.DrawLines(mypen_e, cf);
                    pos = i + 1;
                }

            }
            gfs_e.DrawImage(image, 0, 0, panel_Result_e.Width, panel_Result_e.Height);


            //int index = 0;
            //int code = m_BenchHandle.m_AccessHandle.GetIndexBaseKeyWord("TestResult_1", m_RecordId, ref index);
            //if (code != 1)
            //{
            //    MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);

            //object Obj = default(object);
            //code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "StartTime", ref Obj);
            //if (code != 1)
            //{
            //    return;
            //}
            startTime = m_HistoryStartTime;

            //code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "EndTime", ref Obj);
            //if (code != 1)
            //{
            //    return;
            //}
            endTime = m_HistoryEndTime;

            TimeSpan tp = endTime - startTime;
            int second = (int)tp.TotalSeconds;
            label_X1_e.Text = startTime.ToString("HH:mm:ss");
            textBox_StartTime_e.Text = label_X1_e.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7_e.Text = endTime.ToString("HH:mm:ss");
            textBox_EndTime_e.Text = label_X7_e.Text;

            UInt32 t1 = m_yMax * 1 / 5;
            label45.Text = t1.ToString() + "MPa";

            UInt32 t2 = m_yMax * 2 / 5;
            label44.Text = t2.ToString() + "MPa";

            UInt32 t3 = m_yMax * 3 / 5;
            label43.Text = t3.ToString() + "MPa";

            UInt32 t4 = m_yMax * 4 / 5;
            label41.Text = t4.ToString() + "MPa";

            //textBox_yMax.Text = m_yMax.ToString();
            //textBox_yMax_e.Text = m_yMax.ToString();
        }

        /// <summary>
        /// 读取历史信息
        /// </summary>
        /// <returns></returns>
        private bool ReadCurveDataFromDataBase_e()
        {
            int code = 0;
            int index = 0;
            code = m_BenchHandle.m_AccessHandle.GetIndexBaseKeyWord("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            float[] X;
            float[] Y;

            code = m_BenchHandle.m_AccessHandle.ReadFloatsData("TestResult_1", index, "CuverData_X", out X);
            if (code != 1)
            {
                return false;
            }

            code = m_BenchHandle.m_AccessHandle.ReadFloatsData("TestResult_1", index, "CuverData_Y", out Y);
            if (code != 1)
            {
                return false;
            }

            PointF[] p = new PointF[X.Length];
            m_HistoryPointF = new PointF[X.Length];
            int len = X.Length;
            for (int i = 0; i < len; i++)
            {
                p[i].X = X[i];
                p[i].Y = Y[i];
                m_HistoryPointF[i].X = X[i];
                m_HistoryPointF[i].Y = Y[i];
            }

            //绘制曲线
            //m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result_e);
            //Graphics gfs_e = panel_Result.CreateGraphics();
            //Pen mypen_e = new Pen(Color.Red, 1.5f);

            //int pos = 0;
            //for (int i = 0; i < len; i++)
            //{
            //    if (i != 0 && p[i].X == 0 && p[i].Y == 0)
            //    {
            //        PointF[] f = new PointF[i - pos];
            //        for (int j = pos; j < i; j++)
            //        {
            //            f[j - pos] = p[j];
            //        }
            //        gfs_e.DrawLines(mypen_e, f);
            //        pos = i + 1;
            //    }
            //}

            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);

            object Obj = default(object);
            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "StartTime", ref Obj);
            if (code != 1)
            {
                return false;
            }
            startTime = Convert.ToDateTime(Obj);
            m_HistoryStartTime = startTime;

            code = m_BenchHandle.m_AccessHandle.ReadSigleData("TestResult_1", index, "EndTime", ref Obj);
            if (code != 1)
            {
                return false;
            }
            endTime = Convert.ToDateTime(Obj);
            m_HistoryEndTime = endTime;
            return true;
        }

        private bool ReadCurveDataFromDataBase_e_M()
        {
            int code = 0;
            int index = 0;
            code = m_BenchHandle.m_AccessHandle.GetIndexBaseKeyWord_M("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            float[] X;
            float[] Y;

            code = m_BenchHandle.m_AccessHandle.ReadFloatsData_M("TestResult_1", index, "CuverData_X", out X);
            if (code != 1)
            {
                return false;
            }

            code = m_BenchHandle.m_AccessHandle.ReadFloatsData_M("TestResult_1", index, "CuverData_Y", out Y);
            if (code != 1)
            {
                return false;
            }

            PointF[] p = new PointF[X.Length];
            m_HistoryPointF = new PointF[X.Length];
            int len = X.Length;
            for (int i = 0; i < len; i++)
            {
                p[i].X = X[i];
                p[i].Y = Y[i];
                m_HistoryPointF[i].X = X[i];
                m_HistoryPointF[i].Y = Y[i];
            }

            //绘制曲线
            //m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result_e);
            //Graphics gfs_e = panel_Result.CreateGraphics();
            //Pen mypen_e = new Pen(Color.Red, 1.5f);

            //int pos = 0;
            //for (int i = 0; i < len; i++)
            //{
            //    if (i != 0 && p[i].X == 0 && p[i].Y == 0)
            //    {
            //        PointF[] f = new PointF[i - pos];
            //        for (int j = pos; j < i; j++)
            //        {
            //            f[j - pos] = p[j];
            //        }
            //        gfs_e.DrawLines(mypen_e, f);
            //        pos = i + 1;
            //    }
            //}

            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);

            object Obj = default(object);
            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "StartTime", ref Obj);
            if (code != 1)
            {
                return false;
            }
            startTime = Convert.ToDateTime(Obj);
            m_HistoryStartTime = startTime;

            code = m_BenchHandle.m_AccessHandle.ReadSigleData_M("TestResult_1", index, "EndTime", ref Obj);
            if (code != 1)
            {
                return false;
            }
            endTime = Convert.ToDateTime(Obj);
            m_HistoryEndTime = endTime;
            return true;
        }

        /// <summary>
        /// 将英文产品信息存入到本地缓存
        /// </summary>
        private void ReadTextBoxData_e()
        {
            I_ProductName = TB_Name_e.Text;
            I_ProductNo = TB_ProductID_e.Text;
            I_ProductModel = TB_Module_e.Text;
            I_TestStandard = TB_Standard_e.Text;
            I_TestNiuju = TB_Niuju_e.Text;
            I_EquipmentNo = TB_EquipNum_e.Text;
            I_CompanyName = TB_CompanyName_e.Text;
            I_Operator = TB_Operator_e.Text;
            I_Tester = TB_Tester_e.Text;
            I_Confirmation = TB_Conformation_e.Text;

            WritePrivateProfileString(strSec, "ProductName", I_ProductName, strFilePath);
            WritePrivateProfileString(strSec, "ProductNo", I_ProductNo, strFilePath);
            WritePrivateProfileString(strSec, "ProductModel", I_ProductModel, strFilePath);
            WritePrivateProfileString(strSec, "TestStandard", I_TestStandard, strFilePath);
            WritePrivateProfileString(strSec, "TestNiuju", I_TestNiuju, strFilePath);
            WritePrivateProfileString(strSec, "EquipmentNo", I_EquipmentNo, strFilePath);
            WritePrivateProfileString(strSec, "CompanyName", I_CompanyName, strFilePath);
            WritePrivateProfileString(strSec, "Operator", I_Operator, strFilePath);
            WritePrivateProfileString(strSec, "Tester", I_Tester, strFilePath);
            WritePrivateProfileString(strSec, "Confirmation", I_Confirmation, strFilePath);
            WritePrivateProfileString(strSec, "Date", I_Date, strFilePath);
        }

        /// <summary>
        /// 创建英文报表
        /// </summary>
        /// <returns></returns>
        private bool CreatOfficeReport_e()
        {
            string Name = "EReport" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".doc";
            string FileName = m_BaseFilePath + "\\" + Name;
            //FileName = @"d:\123.doc";
            //创建Word文档
            Object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document WordDoc = WordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            //设置格式
            WordApp.Selection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceAtLeast;//1.5倍行距
            WordApp.Selection.ParagraphFormat.SpaceBefore = float.Parse("0");
            WordApp.Selection.ParagraphFormat.SpaceBeforeAuto = 0;
            WordApp.Selection.ParagraphFormat.SpaceAfter = float.Parse("0");//段后间距
            WordApp.Selection.ParagraphFormat.SpaceAfterAuto = 0;

            WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
            WordDoc.Paragraphs.Last.Range.Font.Size = 9;
            //WordDoc.Paragraphs.Last.Range           
            string strContent = m_OfficeInfo + "\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;
            for (int i = 0; i < D_delaytime; i++) ;

            WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
            WordDoc.Paragraphs.Last.Range.Font.Size = 11;
            //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            strContent = "Hydrostatic Test Chart\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;

            WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
            WordDoc.Paragraphs.Last.Range.Font.Size = 10;
            WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;//水平居中
            strContent = "Product Information\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;

            //添加表格
            Microsoft.Office.Interop.Word.Table table = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, 2, 7, ref oMissing, ref oMissing);
            //设置表格样式
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 8;
            table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
            table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

            table.Cell(1, 1).Range.Text = "Draw No";
            table.Cell(1, 2).Range.Text = "Serial No";
            table.Cell(1, 3).Range.Text = "Test Standard";
            table.Cell(1, 4).Range.Text = "Test Torque(N.M)";
            table.Cell(1, 5).Range.Text = "Equipment No";
            table.Cell(1, 6).Range.Text = "Company Name";
            table.Cell(1, 7).Range.Text = "Contract No";

            table.Cell(2, 1).Range.Text = TB_Name.Text;
            table.Cell(2, 2).Range.Text = TB_ProductID.Text;
            table.Cell(2, 3).Range.Text = TB_Module.Text;
            table.Cell(2, 4).Range.Text = TB_Standard.Text;
            table.Cell(2, 5).Range.Text = TB_Niuju.Text;
            table.Cell(2, 6).Range.Text = TB_EquipNum.Text;
            table.Cell(2, 7).Range.Text = TB_CompanyName.Text;

            //插入文本，试验数据
            strContent = "\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;

            WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
            WordDoc.Paragraphs.Last.Range.Font.Size = 9;
            //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            strContent = "Test Parameter\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;

            //第二个表格
            int rows = dataGridView_DateList.Rows.Count;
            table = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, rows - 1, 5, ref oMissing, ref oMissing);
            //表格样式
            table.Range.Font.Size = 8;
            table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
            table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;



            //for (int i = 1; i < rows; i++)
            //{
            //    for (int j = 0; j <= 5; j++)
            //    {
            //        table.Cell(i, j).Range.Text = dataGridView_DateList[i, j].Value.ToString();
            //    }
            //}

            for (int i = 1; i < rows; i++)
            {
                for (int j = 1; j <= 5; j++)
                {
                    table.Cell(i, j).Range.Text = dataGridView_DateList[j - 1, i - 1].Value.ToString();
                    //for (int t = 0; t < D_delaytime; t++) ;
                }
            }
            table.Cell(1, 1).Range.Text = "Stage No";
            table.Cell(1, 2).Range.Text = "Start Pressure(MPa)";
            table.Cell(1, 3).Range.Text = "Final Pressure(MPa)";
            table.Cell(1, 4).Range.Text = "Hold Period(Min)";
            table.Cell(1, 5).Range.Text = "Pressure Reduction(MPa)";



            strContent = "\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;

            //插入图片
            FileName = m_BaseFilePath + "\\Presure.png";//图片所在路径
            object LinkToFile = false;
            object SaveWithDocument = true;
            object Anchor = WordDoc.Paragraphs.Last.Range;
            //object Anchor = WordDoc.Application.Selection.Range;
            //WordDoc.Paragraphs.Last.Range = WordDoc.Application.Selection.Range;

            WordDoc.Application.ActiveDocument.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Anchor);
            Microsoft.Office.Interop.Word.Shape s = WordDoc.Application.ActiveDocument.InlineShapes[1].ConvertToShape();
            s.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapInline;
            object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;

            object oEndOfDoc = "\\endofdoc";
            Microsoft.Office.Interop.Word.Paragraph return_pragraph;
            object myrange2 = WordDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            return_pragraph = WordDoc.Content.Paragraphs.Add(ref myrange2);
            return_pragraph.Range.InsertParagraphAfter(); //插入一个空白行  


            //插入试验结果
            WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
            WordDoc.Paragraphs.Last.Range.Font.Size = 9;
            //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            strContent = "Test Result：no visibale leakage during each holding period\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;

            //第三个表格
            table = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, 2, 4, ref oMissing, ref oMissing);
            //表格样式
            table.Range.Font.Size = 8;
            table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
            table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

            table.Cell(1, 1).Range.Text = "Tested By";
            table.Cell(1, 2).Range.Text = "Inspected by";
            table.Cell(1, 3).Range.Text = "Review By";
            table.Cell(1, 4).Range.Text = "Date";
            table.Cell(2, 1).Range.Text = TB_Operator.Text;
            table.Cell(2, 2).Range.Text = TB_Tester.Text;
            table.Cell(2, 3).Range.Text = TB_Conformation.Text;
            table.Cell(2, 4).Range.Text = TB_Date.Text;


            Name = "EReport" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".doc";
            object fileName = m_BaseFilePath + "\\Report\\" + Name;

            WordDoc.SaveAs(ref fileName, ref format, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            //关闭wordDoc 文档对象
            WordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
            //关闭wordApp 组件对象
            WordApp.Quit(ref oMissing, ref oMissing, ref oMissing);

            if (File.Exists(fileName.ToString()))
            {
                //File.Open(FileName,FileMode.Open);
                System.Diagnostics.Process.Start(fileName.ToString());
            }

            return true;
        }

        /// <summary>
        /// 英文报表中的返回
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
            //this.Dispose();
        }


        /// <summary>
        /// 英文报表中的创建报表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            bool ret = false;
            SavePic();
            ReadTextBoxData_e();
            if (!m_isOpenDataBase)
            {
                if (!m_isCreatReport)
                {
                    ret = SaveDataToDataBase_M();
                    if (!ret)
                    {
                        MessageBox.Show("创建报表失败,请检查数据信息", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Cursor = System.Windows.Forms.Cursors.Arrow;
                        return;
                    }
                    m_isCreatReport = true;
                }
            }
            ret = CreatOfficeReport_e();
            if (!ret)
            {
                MessageBox.Show("创建报表失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            //m_DrawCurve_e.Start();
        }


        private void button3_Click(object sender, EventArgs e)
        {
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            m_SelectHistoryData_No1Handle = new SelectHistoryData_No1(this, m_BenchHandle);
            m_SelectHistoryData_No1Handle.ControlBox = false;
            m_SelectHistoryData_No1Handle.ShowDialog();


            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            if (m_isNoSelect)
            {
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            bool ret = ReadInfoDataFromDataBase_M();
            if (!ret)
            {
                MessageBox.Show("读取测试信息失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            ret = ReadTestDataFromDataBase_M();
            if (!ret)
            {
                MessageBox.Show("读取测试数据失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            ret = ReadCurveDataFromDataBase_e_M();
            if (!ret)
            {
                MessageBox.Show("读取曲线信息失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }
            m_isOpenDataBase = true;
            //button_CreatReport.Enabled = false;
            //button2.Enabled = false;
            //m_DrawHistoryCurveTimer_e.Start();
            DrawHistoryCurve_e();
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
        }

        #endregion

        /// <summary>
        /// 读取试验曲线, 使用系统DLL，目前不适用
        /// </summary>
        public void ReadTestPic()
        {
            int count = m_BenchHandle.m_PointFLists.Count;
            int len = 0;
            for (int i = 0; i < count; i++)
            {
                int length = m_BenchHandle.m_PointFLists[i].m_pts.Length;
                len += length;
            }
            PointF[] p = new PointF[len];
            if (len <= 0)
            {
                return;
            }
            PointF[] f = new PointF[len - 1];
            int pCount = 0;
            for (int i = 0; i < count; i++)
            {
                int length = m_BenchHandle.m_PointFLists[i].m_pts.Length;
                for (int j = 0; j < length; j++)
                {
                    if (i == 0)
                    {
                        p[j] = m_BenchHandle.m_PointFLists[i].m_pts[j];
                    }
                    else
                    {
                        p[pCount + j].Y = m_BenchHandle.m_PointFLists[i].m_pts[j].Y;
                        p[pCount + j].X = p[pCount - 1].X + m_BenchHandle.m_PointFLists[i].m_pts[j].X;
                    }
                }
                pCount += length;
            }
            for (int i = 0; i < len - 1; i++)
            {
                f[i] = p[i];
            }
            m_BenchHandle.m_DrawCurveHandle.DrawLine(f, panel_Result, Color.Red, p[len - 1].X + 10, m_yMax);
        }

        private void TableControl_English_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (TableControl_English.SelectedIndex == 0)
            {
                if (m_isOpenDataBase)
                {
                    //m_DrawHistoryCurveTimer.Start();
                    DrawHistoryCurve();
                }
                else
                {
                    //m_DrawCurve.Start();
                    ReadTestPic1();
                }
            }

            if (TableControl_English.SelectedIndex == 1)
            {
                if (m_isOpenDataBase)
                {
                    //m_DrawHistoryCurveTimer_e.Start();
                    DrawHistoryCurve_e();
                }
                else
                {
                    //m_DrawCurve_e.Start();
                    ReadTestPic1_e();
                }
            }
            SavePic();
        }

        private string m_OfficeInfo = "";//office信息
        private void ReadOfficeInfo()
        {
            string strFilePath1 = System.Windows.Forms.Application.StartupPath + @"\OfficeInfo.ini";//获取INI文件路径
            string strSecfileName = Path.GetFileNameWithoutExtension(strFilePath1);
            string doc = ContentValue1(strSecfileName, "Doc No", strFilePath1);
            m_OfficeInfo += "Doc No:" + doc + ", ";
            string ed = ContentValue1(strSecfileName, "ED", strFilePath1);
            m_OfficeInfo += "ED:" + ed + ", ";
            string date = ContentValue1(strSecfileName, "Date", strFilePath1);
            m_OfficeInfo += "Date:" + date + ", ";
            string approve = ContentValue1(strSecfileName, "Approved by", strFilePath1);
            m_OfficeInfo += "Approved by:" + approve + ".";
        }

        private string ContentValue1(string Section, string key, string path)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, path);
            return temp.ToString();
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
            textBox_yMax_e.Text = m_yMax.ToString();
            if (m_isOpenDataBase)
            {
                //m_DrawHistoryCurveTimer.Start();
                DrawHistoryCurve();
            }
            else
            {
                //m_DrawCurve.Start();
                ReadTestPic1();
            }
            SavePic();
        }

        private void textBox_yMax_e_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m_yMax = Convert.ToUInt32(textBox_yMax_e.Text);
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
            textBox_yMax_e.Text = m_yMax.ToString();
            if (m_isOpenDataBase)
            {

                //m_DrawHistoryCurveTimer_e.Start();
                DrawHistoryCurve_e();
            }
            else
            {
                //m_DrawCurve_e.Start();
                ReadTestPic1_e();
            }
            SavePic();
        }

        private void GetZoomPointF(PointF[] p, out PointF[] f, int starttime, int endTime)
        {
            int StartPos = 0;
            int EndPos = 0;
            int len = p.Length;
            bool isFindStart = false;
            bool isFindEnd = false;

            for (int i = 0; i < len; i++)
            {
                if (p[i].X == 0 && p[i].Y == 0)
                {
                    continue;
                }
                if (!isFindStart)
                {
                    if (p[i].X >= starttime)
                    {
                        StartPos = i;
                        isFindStart = true;
                    }
                }
                if (!isFindEnd)
                {
                    if (p[i].X >= endTime)
                    {
                        EndPos = i;
                        isFindEnd = true;
                        break;
                    }
                }
            }
            if (!isFindEnd && EndPos == 0)
            {
                EndPos = len;
            }
            if (EndPos == StartPos)
            {
                f = new PointF[1];
                return;
            }
            f = new PointF[EndPos - StartPos];
            for (int i = StartPos; i < EndPos - 1; i++)
            {
                if (p[i].X == 0 && p[i].Y == 0)
                {
                    f[i - StartPos] = p[i];
                    continue;
                }
                else
                {
                    f[i - StartPos].X = p[i].X - starttime;
                    f[i - StartPos].Y = p[i].Y;
                }
            }
            f[EndPos - StartPos - 1].X = 0;
            f[EndPos - StartPos - 1].Y = 0;
        }

        private void DrawZoomPic(PointF[] p, Panel panel, int xMax, float yMax)
        {
            int len = p.Length;
            int pos = 0;

            Graphics gfs = panel.CreateGraphics();
            Pen mypen = new Pen(Color.Red, 1.5f);
            gfs.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            for (int i = 0; i < len; i++)
            {
                if (i != 0 && p[i].X == 0 && p[i].Y == 0)
                {
                    PointF[] f = new PointF[i - pos];
                    for (int j = pos; j < i; j++)
                    {
                        f[j - pos].X = p[j].X;
                        f[j - pos].Y = p[j].Y;
                    }
                    PointF[] cf;
                    PointTranslate(f, out cf, panel, xMax, yMax);
                    for (int m = 1; m < cf.Length; m++)
                    {
                        gfs.DrawLine(mypen, cf[m - 1], cf[m]);
                    }
                    pos = i + 1;
                }

            }
        }

        private void textBox_StartTime_KeyDown(object sender, KeyEventArgs e)
        {
            int keyCode = e.KeyValue;
            if (keyCode != 13)
            {
                return;
            }

            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);
            Graphics gfs = panel_Result.CreateGraphics();
            Pen mypen = new Pen(Color.Red, 1.5f);
            try
            {
                if (m_isOpenDataBase)
                {
                    startTime = Convert.ToDateTime(m_HistoryStartTime.ToShortDateString() + " " + textBox_StartTime.Text);
                    endTime = Convert.ToDateTime(m_HistoryEndTime.ToShortDateString() + " " + textBox_EndTime.Text);
                }
                else
                {
                    startTime = Convert.ToDateTime(textBox_StartTime.Text);
                    endTime = Convert.ToDateTime(textBox_EndTime.Text);
                }
            }
            catch (Exception)
            {
                startTime = m_BenchHandle.m_StartTime;
                //endTime = m_BenchHandle.m_EndTime;
            }
            TimeSpan start_tp = default(TimeSpan);
            TimeSpan end_tp = default(TimeSpan);
            int start_totalTime = 0;
            int end_totalTime = 0;
            if (m_isOpenDataBase)
            {
                start_tp = startTime - m_HistoryStartTime;
                end_tp = endTime - m_HistoryStartTime;
                start_totalTime = (startTime.Hour - m_HistoryStartTime.Hour) * 3600 +
                                    (startTime.Minute - m_HistoryStartTime.Minute) * 60 +
                                    (startTime.Second - m_HistoryStartTime.Second);
                end_totalTime = (endTime.Hour - m_HistoryStartTime.Hour) * 3600 +
                                    (endTime.Minute - m_HistoryStartTime.Minute) * 60 +
                                    (endTime.Second - m_HistoryStartTime.Second);
                if (start_totalTime <= 0)
                {
                    startTime = m_HistoryStartTime;
                    start_totalTime = 0;
                    textBox_StartTime.Text = startTime.ToString("HH:mm:ss");
                }

                PointF[] p;
                GetZoomPointF(m_HistoryPointF, out p, start_totalTime, end_totalTime);

                Graphics gfs_c = panel_Result.CreateGraphics();
                gfs_c.Clear(Color.White);
                gfs_c.Dispose();
                m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result);

                DrawZoomPic(p, panel_Result, end_totalTime - start_totalTime + 2, m_yMax);
            }
            else
            {
                start_tp = startTime - m_BenchHandle.m_StartTime;
                end_tp = endTime - m_BenchHandle.m_StartTime;
                start_totalTime = start_tp.Hours * 3600 + start_tp.Minutes * 60 + start_tp.Seconds;
                end_totalTime = end_tp.Hours * 3600 + end_tp.Minutes * 60 + end_tp.Seconds;
                if (start_totalTime <= 0)
                {
                    startTime = m_BenchHandle.m_StartTime;
                    start_totalTime = 0;
                    textBox_StartTime.Text = startTime.ToString("HH:mm:ss");
                }


                PointF[] p;
                GetZoomPointF(m_ChangedPointF, out p, start_totalTime, end_totalTime);

                Graphics gfs_c = panel_Result.CreateGraphics();
                gfs_c.Clear(Color.White);
                gfs_c.Dispose();
                m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result);

                DrawZoomPic(p, panel_Result, end_totalTime - start_totalTime + 2, m_yMax);
            }
            RefreshTimeLabel(startTime, endTime);
        }

        private void textBox_EndTime_KeyDown(object sender, KeyEventArgs e)
        {
            int keyCode = e.KeyValue;
            if (keyCode != 13)
            {
                return;
            }
            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);
            Graphics gfs = panel_Result.CreateGraphics();
            Pen mypen = new Pen(Color.Red, 1.5f);
            try
            {
                {
                    if (m_isOpenDataBase)
                    {
                        startTime = Convert.ToDateTime(m_HistoryStartTime.ToShortDateString() + " " + textBox_StartTime.Text);
                        endTime = Convert.ToDateTime(m_HistoryEndTime.ToShortDateString() + " " + textBox_EndTime.Text);
                    }
                    else
                    {
                        startTime = Convert.ToDateTime(textBox_StartTime.Text);
                        endTime = Convert.ToDateTime(textBox_EndTime.Text);
                    }
                }
            }
            catch (Exception)
            {
                //startTime = m_BenchHandle.m_StartTime;
                endTime = m_BenchHandle.m_EndTime;
            }
            TimeSpan start_tp = default(TimeSpan);
            TimeSpan end_tp = default(TimeSpan);
            int start_totalTime = 0;
            int end_totalTime = 0;
            if (m_isOpenDataBase)
            {
                start_tp = startTime - m_HistoryStartTime;
                end_tp = endTime - m_HistoryStartTime;
                start_totalTime = (startTime.Hour - m_HistoryStartTime.Hour) * 3600 +
                                    (startTime.Minute - m_HistoryStartTime.Minute) * 60 +
                                    (startTime.Second - m_HistoryStartTime.Second);
                end_totalTime = (endTime.Hour - m_HistoryStartTime.Hour) * 3600 +
                                    (endTime.Minute - m_HistoryStartTime.Minute) * 60 +
                                    (endTime.Second - m_HistoryStartTime.Second);
                int his_totalTime = (endTime.Hour - m_HistoryEndTime.Hour) * 3600 +
                                    (endTime.Minute - m_HistoryEndTime.Minute) * 60 +
                                    (endTime.Second - m_HistoryEndTime.Second);
                if (his_totalTime >= 0)
                {
                    endTime = m_HistoryEndTime;
                    end_totalTime = (endTime.Hour - m_HistoryStartTime.Hour) * 3600 +
                                    (endTime.Minute - m_HistoryStartTime.Minute) * 60 +
                                    (endTime.Second - m_HistoryStartTime.Second);
                    textBox_EndTime.Text = endTime.ToString("HH:mm:ss");
                }

                PointF[] p;
                GetZoomPointF(m_HistoryPointF, out p, start_totalTime, end_totalTime);

                Graphics gfs_c = panel_Result.CreateGraphics();
                gfs_c.Clear(Color.White);
                gfs_c.Dispose();
                m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result);

                DrawZoomPic(p, panel_Result, end_totalTime - start_totalTime + 2, m_yMax);
            }
            else
            {
                DateTime sys_EndTime = default(DateTime);
                m_BenchHandle.ReadEndTime(ref sys_EndTime);
                start_tp = startTime - m_BenchHandle.m_StartTime;
                end_tp = endTime - m_BenchHandle.m_StartTime;
                start_totalTime = start_tp.Hours * 3600 + start_tp.Minutes * 60 + start_tp.Seconds;
                end_totalTime = end_tp.Hours * 3600 + end_tp.Minutes * 60 + end_tp.Seconds;
                int totalTime = (endTime.Hour - sys_EndTime.Hour) * 3600 +
                                    (endTime.Minute - sys_EndTime.Minute) * 60 +
                                    (endTime.Second - sys_EndTime.Second);
                if (totalTime >= 0)
                {
                    endTime = m_BenchHandle.m_EndTime;
                    end_totalTime = (endTime.Hour - m_BenchHandle.m_StartTime.Hour) * 3600 +
                                    (endTime.Minute - m_BenchHandle.m_StartTime.Minute) * 60 +
                                    (endTime.Second - m_BenchHandle.m_StartTime.Second); ;
                    textBox_EndTime.Text = endTime.ToString("HH:mm:ss");
                }


                PointF[] p;
                GetZoomPointF(m_ChangedPointF, out p, start_totalTime, end_totalTime);

                Graphics gfs_c = panel_Result.CreateGraphics();
                gfs_c.Clear(Color.White);
                gfs_c.Dispose();
                m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result);

                DrawZoomPic(p, panel_Result, end_totalTime - start_totalTime + 2, m_yMax);
            }
            RefreshTimeLabel(startTime, endTime);
        }

        private void textBox_StartTime_e_KeyDown(object sender, KeyEventArgs e)
        {
            int keyCode = e.KeyValue;
            if (keyCode != 13)
            {
                return;
            }
            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);
            Graphics gfs = panel_Result_e.CreateGraphics();
            Pen mypen = new Pen(Color.Red, 1.5f);
            try
            {
                if (m_isOpenDataBase)
                {
                    startTime = Convert.ToDateTime(m_HistoryStartTime.ToShortDateString() + " " + textBox_StartTime_e.Text);
                    endTime = Convert.ToDateTime(m_HistoryEndTime.ToShortDateString() + " " + textBox_EndTime_e.Text);
                }
                else
                {
                    startTime = Convert.ToDateTime(textBox_StartTime_e.Text);
                    endTime = Convert.ToDateTime(textBox_EndTime_e.Text);
                }
            }
            catch (Exception)
            {
                startTime = m_BenchHandle.m_StartTime;
                //endTime = m_BenchHandle.m_EndTime;
            }
            TimeSpan start_tp = default(TimeSpan);
            TimeSpan end_tp = default(TimeSpan);
            int start_totalTime = 0;
            int end_totalTime = 0;
            if (m_isOpenDataBase)
            {
                start_tp = startTime - m_HistoryStartTime;
                end_tp = endTime - m_HistoryStartTime;
                start_totalTime = (startTime.Hour - m_HistoryStartTime.Hour) * 3600 +
                                    (startTime.Minute - m_HistoryStartTime.Minute) * 60 +
                                    (startTime.Second - m_HistoryStartTime.Second);
                end_totalTime = (endTime.Hour - m_HistoryStartTime.Hour) * 3600 +
                                    (endTime.Minute - m_HistoryStartTime.Minute) * 60 +
                                    (endTime.Second - m_HistoryStartTime.Second);
                if (start_totalTime <= 0)
                {
                    startTime = m_HistoryStartTime;
                    start_totalTime = 0;
                    textBox_StartTime_e.Text = startTime.ToString("HH:mm:ss");
                }

                PointF[] p;
                GetZoomPointF(m_HistoryPointF, out p, start_totalTime, end_totalTime);

                Graphics gfs_c = panel_Result_e.CreateGraphics();
                gfs_c.Clear(Color.White);
                gfs_c.Dispose();
                m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result_e);

                DrawZoomPic(p, panel_Result_e, end_totalTime - start_totalTime + 2, m_yMax);
            }
            else
            {
                start_tp = startTime - m_BenchHandle.m_StartTime;
                end_tp = endTime - m_BenchHandle.m_StartTime;
                start_totalTime = start_tp.Hours * 3600 + start_tp.Minutes * 60 + start_tp.Seconds;
                end_totalTime = end_tp.Hours * 3600 + end_tp.Minutes * 60 + end_tp.Seconds;
                if (start_totalTime <= 0)
                {
                    startTime = m_BenchHandle.m_StartTime;
                    start_totalTime = 0;
                    textBox_StartTime_e.Text = startTime.ToString("HH:mm:ss");
                }


                PointF[] p;
                GetZoomPointF(m_ChangedPointF, out p, start_totalTime, end_totalTime);

                Graphics gfs_c = panel_Result_e.CreateGraphics();
                gfs_c.Clear(Color.White);
                gfs_c.Dispose();
                m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result_e);

                DrawZoomPic(p, panel_Result_e, end_totalTime - start_totalTime + 2, m_yMax);
            }
            RefreshTimeLabel_e(startTime, endTime);
        }

        private void textBox_EndTime_e_KeyDown(object sender, KeyEventArgs e)
        {
            int keyCode = e.KeyValue;
            if (keyCode != 13)
            {
                return;
            }
            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);
            Graphics gfs = panel_Result_e.CreateGraphics();
            Pen mypen = new Pen(Color.Red, 1.5f);
            try
            {
                if (m_isOpenDataBase)
                {
                    startTime = Convert.ToDateTime(m_HistoryStartTime.ToShortDateString() + " " + textBox_StartTime_e.Text);
                    endTime = Convert.ToDateTime(m_HistoryEndTime.ToShortDateString() + " " + textBox_EndTime_e.Text);
                }
                else
                {
                    startTime = Convert.ToDateTime(textBox_StartTime_e.Text);
                    endTime = Convert.ToDateTime(textBox_EndTime_e.Text);
                }
            }
            catch (Exception)
            {
                //startTime = m_BenchHandle.m_StartTime;
                endTime = m_BenchHandle.m_EndTime;
            }
            TimeSpan start_tp = default(TimeSpan);
            TimeSpan end_tp = default(TimeSpan);
            int start_totalTime = 0;
            int end_totalTime = 0;
            if (m_isOpenDataBase)
            {
                start_tp = startTime - m_HistoryStartTime;
                end_tp = endTime - m_HistoryStartTime;
                start_totalTime = (startTime.Hour - m_HistoryStartTime.Hour) * 3600 +
                                    (startTime.Minute - m_HistoryStartTime.Minute) * 60 +
                                    (startTime.Second - m_HistoryStartTime.Second);
                end_totalTime = (endTime.Hour - m_HistoryStartTime.Hour) * 3600 +
                                    (endTime.Minute - m_HistoryStartTime.Minute) * 60 +
                                    (endTime.Second - m_HistoryStartTime.Second);
                int his_totalTime = (endTime.Hour - m_HistoryEndTime.Hour) * 3600 +
                                    (endTime.Minute - m_HistoryEndTime.Minute) * 60 +
                                    (endTime.Second - m_HistoryEndTime.Second);
                if (his_totalTime >= 0)
                {
                    endTime = m_HistoryEndTime;
                    end_totalTime = (endTime.Hour - m_HistoryStartTime.Hour) * 3600 +
                                    (endTime.Minute - m_HistoryStartTime.Minute) * 60 +
                                    (endTime.Second - m_HistoryStartTime.Second);
                    textBox_EndTime_e.Text = endTime.ToString("HH:mm:ss");
                }

                PointF[] p;
                GetZoomPointF(m_HistoryPointF, out p, start_totalTime, end_totalTime);

                Graphics gfs_c = panel_Result_e.CreateGraphics();
                gfs_c.Clear(Color.White);
                gfs_c.Dispose();
                m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result_e);

                DrawZoomPic(p, panel_Result_e, end_totalTime - start_totalTime + 2, m_yMax);
            }
            else
            {
                DateTime sys_EndTime = default(DateTime);
                m_BenchHandle.ReadEndTime(ref sys_EndTime);
                start_tp = startTime - m_BenchHandle.m_StartTime;
                end_tp = endTime - m_BenchHandle.m_StartTime;
                start_totalTime = start_tp.Hours * 3600 + start_tp.Minutes * 60 + start_tp.Seconds;
                end_totalTime = end_tp.Hours * 3600 + end_tp.Minutes * 60 + end_tp.Seconds;
                int totalTime = (endTime.Hour - sys_EndTime.Hour) * 3600 +
                                    (endTime.Minute - sys_EndTime.Minute) * 60 +
                                    (endTime.Second - sys_EndTime.Second);
                if (totalTime >= 0)
                {
                    endTime = m_BenchHandle.m_EndTime;
                    end_totalTime = (endTime.Hour - m_BenchHandle.m_StartTime.Hour) * 3600 +
                                    (endTime.Minute - m_BenchHandle.m_StartTime.Minute) * 60 +
                                    (endTime.Second - m_BenchHandle.m_StartTime.Second);
                    textBox_EndTime_e.Text = endTime.ToString("HH:mm:ss");
                }


                PointF[] p;
                GetZoomPointF(m_ChangedPointF, out p, start_totalTime, end_totalTime);

                Graphics gfs_c = panel_Result_e.CreateGraphics();
                gfs_c.Clear(Color.White);
                gfs_c.Dispose();
                m_BenchHandle.m_DrawCurveHandle.DrawModule(panel_Result_e);

                DrawZoomPic(p, panel_Result_e, end_totalTime - start_totalTime + 2, m_yMax);
            }
            RefreshTimeLabel_e(startTime, endTime);
        }

        private void RefreshTimeLabel(DateTime StartTime, DateTime EndTime)
        {
            TimeSpan tp = EndTime - StartTime;
            int second = (int)tp.TotalSeconds;
            label_X1.Text = StartTime.ToString("HH:mm:ss");
            //textBox_StartTime.Text = label_X1.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7.Text = EndTime.ToString("HH:mm:ss");
            //textBox_EndTime.Text = label_X7.Text;
        }

        private void RefreshTimeLabel_e(DateTime StartTime, DateTime EndTime)
        {
            TimeSpan tp = EndTime - StartTime;
            int second = (int)tp.TotalSeconds;
            label_X1_e.Text = StartTime.ToString("HH:mm:ss");
            //textBox_StartTime_e.Text = label_X1_e.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7_e.Text = EndTime.ToString("HH:mm:ss");
            //textBox_EndTime_e.Text = label_X7_e.Text;
        }

    }
}
