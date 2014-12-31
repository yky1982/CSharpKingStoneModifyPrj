using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Runtime.InteropServices;
using XMLConfigPLC200SampleData;

namespace KingStoneModify
{
    public partial class Form1 : Form
    {
        #region com
        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "test", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void test(ref int RevCount, ref int SendCount, ref int SendFailCount, ref int frameNumfail);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "test2", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void test2(ref int Count, ref int d300ms, ref int d1s, ref int d5s);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "InterfaceInit", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int InterfaceInit();

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "Dispose", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int Dispose();

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "XMLLoad", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int XMLLoad(string FileName);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "Connect", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int Connect(string GuestIP, int Port);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "Start", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int Start();

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "Stop", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int Stop();

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "Suspend", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int Suspend();

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "Resume", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int Resume();

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref bool data, ref DateTime dt);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref byte data, ref DateTime dt);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref sbyte data, ref DateTime dt);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref Int16 data, ref DateTime dt);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref UInt16 data, ref DateTime dt);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref Int32 data, ref DateTime dt);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref UInt32 data, ref DateTime dt);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref float data, ref DateTime dt);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadOneData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadOneData(string Regname, ref bool data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadOneData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadOneData(string Regname, ref byte data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadOneData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadOneData(string Regname, ref sbyte data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadOneData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadOneData(string Regname, ref Int16 data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadOneData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadOneData(string Regname, ref UInt16 data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadOneData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadOneData(string Regname, ref Int32 data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadOneData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadOneData(string Regname, ref UInt32 data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadOneData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadOneData(string Regname, ref float data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReadDataInfo", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadDataInfo(string Regname, ref bool RW, ref string DataType);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, bool Data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, byte Data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, sbyte Data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, Int16 Data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, UInt16 Data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, Int32 Data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, UInt32 Data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, float Data);

        [DllImport("XMLConfigPLC200SampleData.dll", EntryPoint = "ReportComState", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReportComState(ref bool state);

        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        private Form[] m_FromHandle;

        public DLLInf m_PLCCommHandle = new DLLInf();
        public string m_PLCIPAddress = @"192.168.1.188";
        public int m_PLCPort = 102;

        private string m_XMLFilePath = Application.StartupPath + "\\SystemFile\\xmlConfig.xml";
        private void Form1_Load(object sender, EventArgs e)
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            PLCComInit();
            FormInit();
            m_PLCCommHandle.Resume();
        }

        private void PLCComInit()
        {
            int code = m_PLCCommHandle.InterfaceInit();
            if (code != 1)
            {
                MessageBox.Show("Plc初始化错误! 错误代码：" + code.ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            code = m_PLCCommHandle.XMLLoad(m_XMLFilePath);
            if (code != 1)
            {
                MessageBox.Show("XML初始化错误! 错误代码：" + code.ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            code = m_PLCCommHandle.Connect(m_PLCIPAddress, m_PLCPort);
            if (code != 1)
            {
                code = m_PLCCommHandle.Connect(m_PLCIPAddress, m_PLCPort);
                if (code != 1)
                {
                    MessageBox.Show("连接错误! 错误代码：" + code.ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            m_PLCCommHandle.Start();
        }

        private void FormInit()
        {
            m_FromHandle = new Form[] { new Bench_1(this, Grp_No1), new Bench_2(this, Grp_No2), new Bench_3(this, Grp_No3), new Bench_4(this, Grp_No4) };
            m_FromHandle[0].TopLevel = false;
            Grp_No1.Controls.Add(m_FromHandle[0]);
            m_FromHandle[0].Show();
            m_FromHandle[1].TopLevel = false;
            Grp_No2.Controls.Add(m_FromHandle[1]);
            m_FromHandle[1].Show();
            m_FromHandle[2].TopLevel = false;
            Grp_No3.Controls.Add(m_FromHandle[2]);
            m_FromHandle[2].Show();
            m_FromHandle[3].TopLevel = false;
            Grp_No4.Controls.Add(m_FromHandle[3]);
            m_FromHandle[3].Show();

        }
    }
}
