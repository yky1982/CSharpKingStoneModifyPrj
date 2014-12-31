using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KingStoneModify
{
    public partial class SelectHistoryData_No1 : Form
    {
        private CreateTestResult_No1 m_ParentHandle;
        private Bench_1 m_BenchHandle;
        public SelectHistoryData_No1(CreateTestResult_No1 handle, Bench_1 benchHandle)
        {
            InitializeComponent();
            m_ParentHandle = handle;
            m_BenchHandle = benchHandle;
        }

        private void SelectHistoryData_No1_Load(object sender, EventArgs e)
        {
            string com = "select * from TestResult_1 ";
            int code = m_BenchHandle.m_AccessHandle.QueryData_M(com, dataGridView_data);
            if (code != 1)
            {
                MessageBox.Show("表达式错误，请验证，错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private int m_RecordId = 0;//数据库中RecordId值，该值唯一

        private void dataGridView_data_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            m_RecordId = Convert.ToInt32(dataGridView_data[0, e.RowIndex].Value.ToString());
        }

        private void button_Sure_Click(object sender, EventArgs e)
        {
            m_ParentHandle.m_RecordId = m_RecordId;
            m_ParentHandle.m_isNoSelect = false;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.Close();
            this.Dispose();
        }

        private void button_Return_Click(object sender, EventArgs e)
        {
            m_ParentHandle.m_isNoSelect = true;
            this.Close();
            this.Dispose();
        }
    }
}
