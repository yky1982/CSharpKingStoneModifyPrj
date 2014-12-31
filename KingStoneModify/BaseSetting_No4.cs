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

namespace KingStoneModify
{
    public partial class BaseSetting_No4 : Form
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);

        private Form1 m_MainFrameHandle;
        private SettingForm_No4 m_SettingFormHandle;
        public BaseSetting_No4(Form1 Handle, SettingForm_No4 handle)
        {
            InitializeComponent();
            m_MainFrameHandle = Handle;
            m_SettingFormHandle = handle;
        }

        private void BaseSetting_No4_Load(object sender, EventArgs e)
        {
            this.ControlBox = false;
            ReadDataFromIni();
        }

        private bool SetPara()
        {
            float HighBumpStartPress = 0;
            float SensorLength = 0;
            float SensorOffSet = 0.0f;
            UInt16 StabilityTime = 0;
            UInt16 OpenValveTime = 0;
            float DropPressSelect = 0;
            UInt16 TestPressInterval = 0;
            //byte SensorAdj = 1;
            try
            {
                HighBumpStartPress = Convert.ToSingle(TB_HighBumpStartPress.Text);
                SensorLength = Convert.ToSingle(TB_SensorLength.Text);
                SensorOffSet = Convert.ToSingle(TB_SensorOffset.Text);
                StabilityTime = Convert.ToUInt16(TB_StabilityTime.Text);
                OpenValveTime = Convert.ToUInt16(TB_OpenValveTime.Text);
                DropPressSelect = Convert.ToSingle(TB_DropPressSelect.Text);
                TestPressInterval = Convert.ToUInt16(TB_TestPressIntervalTime.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("输入的参数错误，请查证。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("HighBumpStartPress_Bench4", HighBumpStartPress);
            if (code != 1)
            {
                MessageBox.Show("高压泵启动压力设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.WriteData("SensorLength_Bench4", SensorLength);
            if (code != 1)
            {
                MessageBox.Show("传感器量程设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.WriteData("SensorOffSet_Bench4", SensorOffSet);
            if (code != 1)
            {
                MessageBox.Show("传感器偏置设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepPressStabilityTime_Bench4", StabilityTime);
            if (code != 1)
            {
                MessageBox.Show("保压稳定时间失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.WriteData("OpenValveTime_Bench4", OpenValveTime);
            if (code != 1)
            {
                MessageBox.Show("开阀操作时间设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.WriteData("DropPressSelect_Bench4", DropPressSelect);
            if (code != 1)
            {
                MessageBox.Show("泄压判断压力设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.WriteData("TestPressInterval_Bench4", TestPressInterval);
            if (code != 1)
            {
                MessageBox.Show("试压间隔时间设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            MessageBox.Show("参数设置成功.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            WritePrivateProfileString(strSec, "HighBumpStartPress_Bench4", HighBumpStartPress.ToString("0.00"), strFilePath);
            WritePrivateProfileString(strSec, "SensorLength_Bench4", SensorLength.ToString("0.00"), strFilePath);
            WritePrivateProfileString(strSec, "SensorOffSet_Bench4", SensorOffSet.ToString("0.00"), strFilePath);
            WritePrivateProfileString(strSec, "KeepPressStabilityTime_Bench4", StabilityTime.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "OpenValveTime_Bench4", OpenValveTime.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "DropPressSelect_Bench4", DropPressSelect.ToString("0.00"), strFilePath);
            WritePrivateProfileString(strSec, "TestPressInterval_Bench4", TestPressInterval.ToString(), strFilePath);
            return true;
        }

        private bool ReadPara()
        {
            float HighBumpStartPress = 0;
            float SensorLength = 0;
            float SensorOffSet = 0.0f;
            UInt16 StabilityTime = 0;
            UInt16 OpenValveTime = 0;
            float DropPressSelect = 0;
            UInt16 TestPressInterval = 0;

            DateTime dt = default(DateTime);

            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("HighBumpStartPress_Bench4", ref HighBumpStartPress, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SensorLength_Bench4", ref SensorLength, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SensorOffSet_Bench4", ref SensorOffSet, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressStabilityTime_Bench4", ref StabilityTime, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("OpenValveTime_Bench4", ref OpenValveTime, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("DropPressSelect_Bench4", ref DropPressSelect, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("TestPressInterval_Bench4", ref TestPressInterval, ref dt);
            if (code != 1)
            {
                return false;
            }

            TB_HighBumpStartPress.Text = HighBumpStartPress.ToString("0.00");
            TB_SensorLength.Text = SensorLength.ToString("0.00");
            TB_SensorOffset.Text = SensorOffSet.ToString("0.00");
            TB_StabilityTime.Text = StabilityTime.ToString();
            TB_OpenValveTime.Text = OpenValveTime.ToString();
            TB_DropPressSelect.Text = DropPressSelect.ToString("0.00");
            TB_TestPressIntervalTime.Text = TestPressInterval.ToString();
            return true;
        }

        private string strFilePath = Application.StartupPath + @"\4#\BaseSettingConfig.ini";//获取INI文件路径
        private string strSec = "BaseSettingConfig"; //INI文件名
        private void ReadDataFromIni()
        {
            TB_HighBumpStartPress.Text = ContentValue(strSec, "HighBumpStartPress_Bench4");
            TB_SensorLength.Text = ContentValue(strSec, "SensorLength_Bench4");
            TB_SensorOffset.Text = ContentValue(strSec, "SensorOffSet_Bench4");
            TB_StabilityTime.Text = ContentValue(strSec, "KeepPressStabilityTime_Bench4");
            TB_OpenValveTime.Text = ContentValue(strSec, "OpenValveTime_Bench4");
            TB_DropPressSelect.Text = ContentValue(strSec, "DropPressSelect_Bench4");
            TB_TestPressIntervalTime.Text = ContentValue(strSec, "TestPressInterval_Bench4");
        }

        private string ContentValue(string Section, string key)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }

        private void button_Read_Click(object sender, EventArgs e)
        {
            bool flag = ReadPara();
            if (!flag)
            {
                MessageBox.Show("参数读取失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button_Set_Click(object sender, EventArgs e)
        {
            bool flag = SetPara();
            if (!flag)
            {
                return;
            }
        }

        private void button_Return_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
