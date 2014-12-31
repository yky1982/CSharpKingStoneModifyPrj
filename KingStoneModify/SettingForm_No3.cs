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
    public partial class SettingForm_No3 : Form
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);

        private Form1 m_MainFrameHandle;
        private Bench_3 m_ParentFormHandle;
        private BaseSetting_No3 m_BaseSettingHandle;
        private string strFilePath = Application.StartupPath + @"\3#\TestConfig.ini";//获取INI文件路径
        private string strSec = "TestConfig"; //INI文件名
        public SettingForm_No3(Form1 Handle, Bench_3 ParentForHandle)
        {
            InitializeComponent();
            m_MainFrameHandle = Handle;
            m_ParentFormHandle = ParentForHandle;
        }

        private void SettingForm_No3_Load(object sender, EventArgs e)
        {
            Load_INI();
            if (Button_Selection_1.Text == "开")
            {
                TextBox_KeepPressure_1.Enabled = true;
                TextBox_KeepTime_1.Enabled = true;
                m_KeepSelect_No1 = true;
            }
            else
            {
                TextBox_KeepPressure_1.Enabled = false;
                TextBox_KeepTime_1.Enabled = false;
                m_KeepSelect_No1 = false;
            }

            if (Button_Selection_2.Text == "开")
            {
                TextBox_KeepPressure_2.Enabled = true;
                TextBox_KeepTime_2.Enabled = true;
                m_KeepSelect_No2 = true;
            }
            else
            {
                TextBox_KeepPressure_2.Enabled = false;
                TextBox_KeepTime_2.Enabled = false;
                m_KeepSelect_No2 = false;
            }

            if (Button_Selection_3.Text == "开")
            {
                TextBox_KeepPressure_3.Enabled = true;
                TextBox_KeepTime_3.Enabled = true;
                m_KeepSelect_No3 = true;
            }
            else
            {
                TextBox_KeepPressure_3.Enabled = false;
                TextBox_KeepTime_3.Enabled = false;
                m_KeepSelect_No3 = false;
            }
        }

        private void Load_INI()
        {
            if (File.Exists(strFilePath))//读取时先要判读INI文件是否存在
            {
                strSec = Path.GetFileNameWithoutExtension(strFilePath);
                TextBox_KeepPressure_1.Text = ContentValue(strSec, "KeepPressure_No1_Bench3");
                TextBox_KeepTime_1.Text = ContentValue(strSec, "KeepTime_No1_Bench3");
                string ButtonState = ContentValue(strSec, "KeepSelect_No1_Bench3");
                if (ButtonState == "OFF")
                {
                    Button_Selection_1.BackColor = Color.Red;
                    Button_Selection_1.Text = "关";

                }
                else
                {
                    Button_Selection_1.BackColor = Color.Green;
                    Button_Selection_1.Text = "开";
                }
                TextBox_KeepPressure_2.Text = ContentValue(strSec, "KeepPressure_No2_Bench3");
                TextBox_KeepTime_2.Text = ContentValue(strSec, "KeepTime_No2_Bench3");
                ButtonState = ContentValue(strSec, "KeepSelect_No2_Bench3");
                if (ButtonState == "OFF")
                {
                    Button_Selection_2.BackColor = Color.Red;
                    Button_Selection_2.Text = "关";
                }
                else
                {
                    Button_Selection_2.BackColor = Color.Green;
                    Button_Selection_2.Text = "开";
                }
                TextBox_KeepPressure_3.Text = ContentValue(strSec, "KeepPressure_No3_Bench3");
                TextBox_KeepTime_3.Text = ContentValue(strSec, "KeepTime_No3_Bench3");
                ButtonState = ContentValue(strSec, "KeepSelect_No3_Bench3");
                if (ButtonState == "OFF")
                {
                    Button_Selection_3.BackColor = Color.Red;
                    Button_Selection_3.Text = "关";
                }
                else
                {
                    Button_Selection_3.BackColor = Color.Green;
                    Button_Selection_3.Text = "开";
                }

                TextBox_Text_No.Text = ContentValue(strSec, "TestNo_Bench3");

            }
            else
            {
                MessageBox.Show("INI文件不存在");
            }
        }

        private string ContentValue(string Section, string key)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }

        private void Button_Return_Click(object sender, EventArgs e)
        {
            this.Hide();
            m_ParentFormHandle.m_TestNo = TextBox_Text_No.Text;
            m_ParentFormHandle.Show();
        }

        private void Button_SavePara_Click(object sender, EventArgs e)
        {
            float Press1 = 0;
            float Time1 = 0;
            string KeepSelect1 = default(string);
            float ContinueTime = 0;
            bool flag = Seriel_No1(ref Press1, ref Time1, ref KeepSelect1);
            if (!flag)
            {
                return;
            }
            float Press2 = 0;
            float Time2 = 0;
            string KeepSelect2 = default(string);
            flag = Seriel_No2(ref Press2, ref Time2, ref KeepSelect2);
            if (!flag)
            {
                return;
            }
            float Press3 = 0;
            float Time3 = 0;
            string KeepSelect3 = default(string);
            flag = Seriel_No3(ref Press3, ref Time3, ref KeepSelect3);
            if (!flag)
            {
                return;
            }

            int code = 1;
            UInt16 setTestTimes = 0;

            if (KeepSelect1 == "OFF")
            {
                bool bSelect1 = false;
                bSelect1 = false;
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No1_Bench3", bSelect1);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压选择1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepPressure_No1_Bench3", Press1);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压压力1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepTime_No1_Bench3", Time1);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压时间1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                bool bSelect1 = true;
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No1_Bench3", bSelect1);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压选择1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ContinueTime += Time1;
                setTestTimes++;
            }

            if (KeepSelect2 == "OFF")
            {
                bool bSelect2 = false;
                bSelect2 = false;
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No2_Bench3", bSelect2);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压选择2写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepPressure_No2_Bench3", Press2);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压压力2写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepTime_No2_Bench3", Time2);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压时间2写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                bool bSelect2 = true;
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No2_Bench3", bSelect2);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压选择2写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ContinueTime += Time2;
                setTestTimes++;

            }

            if (KeepSelect3 == "OFF")
            {
                bool bSelect3 = false;
                bSelect3 = false;
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No3_Bench3", bSelect3);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压选择3写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepPressure_No3_Bench3", Press3);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压压力3写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepTime_No3_Bench3", Time3);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压时间3写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                bool bSelect3 = true;
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No3_Bench3", bSelect3);
                if (code != 1)
                {
                    //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                    MessageBox.Show("保压选择3写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ContinueTime += Time3;
                setTestTimes++;

            }

            WritePrivateProfileString(strSec, "KeepPressure_No1_Bench3", Press1.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "KeepTime_No1_Bench3", Time1.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "KeepSelect_No1_Bench3", KeepSelect1, strFilePath);

            WritePrivateProfileString(strSec, "KeepPressure_No2_Bench3", Press2.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "KeepTime_No2_Bench3", Time2.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "KeepSelect_No2_Bench3", KeepSelect2, strFilePath);

            WritePrivateProfileString(strSec, "KeepPressure_No3_Bench3", Press3.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "KeepTime_No3_Bench3", Time3.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "KeepSelect_No3_Bench3", KeepSelect3, strFilePath);

            WritePrivateProfileString(strSec, "TestNo_Bench3", TextBox_Text_No.Text, strFilePath);
            //如果设置新的实验标号，则清除内存
            string TestNo = TextBox_Text_No.Text;
            if (TestNo != m_ParentFormHandle.m_TestNo)
            {
                m_ParentFormHandle.m_PointFArrays.Clear();
                m_ParentFormHandle.m_PointFLists.Clear();
                m_ParentFormHandle.m_TestResultLists.Clear();
                m_ParentFormHandle.m_TestSequence = 1;
                m_ParentFormHandle.m_isTestStartBaseTestNo = true;
            }
            m_ParentFormHandle.m_TestNo = TestNo;   //保存试验编号

            //m_ParentFormHandle.m_DrawCurveTimeSpan = (int)(ContinueTime * 60);

            MessageBox.Show("写入成功", "Info", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            m_ParentFormHandle.m_isSetFlag = true;
            m_ParentFormHandle.m_SetTestTimes = setTestTimes;
        }

        private bool Seriel_No1(ref float Press, ref float Time, ref string Selection)
        {
            try
            {
                Press = Convert.ToSingle(TextBox_KeepPressure_1.Text);
                Time = Convert.ToSingle(TextBox_KeepTime_1.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("序列1输入错误，请重新检查输入参数", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (Button_Selection_1.Text == "关")
            {
                Selection = "OFF";
            }
            else
            {
                Selection = "ON";
            }

            return true;
        }

        private bool Seriel_No2(ref float Press, ref float Time, ref string Selection)
        {
            try
            {
                Press = Convert.ToSingle(TextBox_KeepPressure_2.Text);
                Time = Convert.ToSingle(TextBox_KeepTime_2.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("序列1输入错误，请重新检查输入参数", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (Button_Selection_2.Text == "关")
            {
                Selection = "OFF";
            }
            else
            {
                Selection = "ON";
            }

            return true;
        }

        private bool Seriel_No3(ref float Press, ref float Time, ref string Selection)
        {
            try
            {
                Press = Convert.ToSingle(TextBox_KeepPressure_3.Text);
                Time = Convert.ToSingle(TextBox_KeepTime_3.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("序列1输入错误，请重新检查输入参数", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (Button_Selection_3.Text == "关")
            {
                Selection = "OFF";
            }
            else
            {
                Selection = "ON";
            }

            return true;
        }

        private bool m_KeepSelect_No1 = true;
        private bool m_KeepSelect_No2 = false;
        private bool m_KeepSelect_No3 = false;

        private void Button_Selection_3_Click(object sender, EventArgs e)
        {
            if (Button_Selection_3.Text == "开")
            {
                if (m_KeepSelect_No2 && m_KeepSelect_No1)
                {
                    TextBox_KeepPressure_3.Enabled = false;
                    TextBox_KeepTime_3.Enabled = false;
                    m_KeepSelect_No3 = false;
                    Button_Selection_3.BackColor = Color.Red;
                    Button_Selection_3.Text = "关";
                }
            }
            else
            {
                if (m_KeepSelect_No2 && m_KeepSelect_No1)
                {
                    TextBox_KeepPressure_3.Enabled = true;
                    TextBox_KeepTime_3.Enabled = true;
                    m_KeepSelect_No3 = true;
                    Button_Selection_3.BackColor = Color.Green;
                    Button_Selection_3.Text = "开";
                }
            }
        }

        private void Button_Selection_2_Click(object sender, EventArgs e)
        {
            if (Button_Selection_2.Text == "开")
            {
                if (m_KeepSelect_No3 == false && m_KeepSelect_No1)
                {
                    TextBox_KeepPressure_2.Enabled = false;
                    TextBox_KeepTime_2.Enabled = false;
                    m_KeepSelect_No2 = false;
                    Button_Selection_2.BackColor = Color.Red;
                    Button_Selection_2.Text = "关";
                }
                else
                {
                    MessageBox.Show("请先关闭序列3", "Info");
                }
            }
            else
            {
                TextBox_KeepPressure_2.Enabled = true;
                TextBox_KeepTime_2.Enabled = true;
                if (m_KeepSelect_No3 == false && m_KeepSelect_No1)
                {
                    m_KeepSelect_No2 = true;
                    Button_Selection_2.BackColor = Color.Green;
                    Button_Selection_2.Text = "开";
                }
            }
        }

        private void Button_Selection_1_Click(object sender, EventArgs e)
        {
            if (Button_Selection_1.Text == "开")
            {
                //TextBox_KeepPressure_1.Enabled = false;
                //TextBox_KeepTime_1.Enabled = false;
                if (m_KeepSelect_No3 == false && m_KeepSelect_No2 == false)
                {
                    MessageBox.Show("不能关闭，必须保留一组测试", "Info");
                    return;
                }
            }
            else
            {
                //TextBox_KeepPressure_1.Enabled = true;
                //TextBox_KeepTime_1.Enabled = true;
            }
        }

        private void Button_BaseSetting_Click(object sender, EventArgs e)
        {
            m_BaseSettingHandle = new BaseSetting_No3(m_MainFrameHandle, this);
            m_BaseSettingHandle.ShowDialog();
            m_BaseSettingHandle.Close();
            m_BaseSettingHandle.Dispose();
            m_BaseSettingHandle = null;
        }

        private void Button_ReadPara_Click(object sender, EventArgs e)
        {
            bool ret = ReadPara();
            if (!ret)
            {
                MessageBox.Show("读取失败，请重试", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private bool ReadPara()
        {
            float KeepPressure1 = 0.0f;
            float KeepPressure2 = 0.0f;
            float KeepPressure3 = 0.0f;
            float KeepTime1 = 0.0f;
            float KeepTime2 = 0.0f;
            float KeepTime3 = 0.0f;
            bool ChanelSelect1 = false;
            bool ChanelSelect2 = false;
            bool ChanelSelect3 = false;

            DateTime dt = default(DateTime);

            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No1_Bench3", ref KeepPressure1, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No2_Bench3", ref KeepPressure2, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No3_Bench3", ref KeepPressure3, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No1_Bench3", ref KeepTime1, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No2_Bench3", ref KeepTime2, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No3_Bench3", ref KeepTime3, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No1_Bench3", ref ChanelSelect1, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No2_Bench3", ref ChanelSelect2, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No3_Bench3", ref ChanelSelect3, ref dt);
            if (code != 1)
            {
                return false;
            }

            if (ChanelSelect1 == true)
            {
                TextBox_KeepPressure_1.Enabled = true;
                TextBox_KeepTime_1.Enabled = true;
                TextBox_KeepPressure_1.Text = KeepPressure1.ToString("0.00");
                TextBox_KeepTime_1.Text = KeepTime1.ToString("0.00");
                Button_Selection_1.Text = "开";
                Button_Selection_1.BackColor = Color.Green;
                m_KeepSelect_No1 = true;
            }
            else
            {
            }

            if (ChanelSelect2 == true)
            {
                TextBox_KeepPressure_2.Enabled = true;
                TextBox_KeepTime_2.Enabled = true;
                TextBox_KeepPressure_2.Text = KeepPressure2.ToString("0.00");
                TextBox_KeepTime_2.Text = KeepTime2.ToString("0.00");
                Button_Selection_2.Text = "开";
                Button_Selection_2.BackColor = Color.Green;
                m_KeepSelect_No2 = true;
            }
            else
            {
                TextBox_KeepPressure_2.Enabled = false;
                TextBox_KeepTime_2.Enabled = false;
                Button_Selection_2.Text = "关";
                Button_Selection_2.BackColor = Color.Red;
                m_KeepSelect_No2 = false;
            }

            if (ChanelSelect3 == true)
            {
                TextBox_KeepPressure_3.Enabled = true;
                TextBox_KeepTime_3.Enabled = true;
                TextBox_KeepPressure_3.Text = KeepPressure3.ToString("0.00");
                TextBox_KeepTime_3.Text = KeepTime3.ToString("0.00");
                Button_Selection_3.Text = "开";
                Button_Selection_3.BackColor = Color.Green;
                m_KeepSelect_No3 = true;
            }
            else
            {
                TextBox_KeepPressure_3.Enabled = false;
                TextBox_KeepTime_3.Enabled = false;
                Button_Selection_3.Text = "关";
                Button_Selection_3.BackColor = Color.Red;
                m_KeepSelect_No3 = false;
            }


            return true;
        }
    }
}
