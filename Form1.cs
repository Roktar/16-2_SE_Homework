using System;
using System.Data;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace adminPrg
{
    public partial class MainForm : Form {

        //DBConnect db = new DBConnect();
        MySqlConnection c = null;
        DataSet ds = new DataSet();
        int num = -1;

        public MainForm()
        {
            InitializeComponent();
            String connector = "Server=; Port=; database=mysql; uid=; pwd=; charset=utf8;";
            c = new MySqlConnection(connector);
            c.Open();

            MySqlDataAdapter cmd = new MySqlDataAdapter(@"SELECT id_num FROM NOTICE;", c);
            cmd.Fill(ds);

            num = ds.Tables[0].Rows.Count;
        }

        private void openform_Click(object sender, EventArgs e)
        {
            searchForm form = new searchForm();
            form.Show();
        }

        private void sendmessage_Click(object sender, EventArgs e)
        {
            if (value.Text == "") {
                MessageBox.Show("공지 내용이 입력되어야 합니다.");
                return;
            }
            string query = @"INSERT INTO NOTICE VALUES(" + num + ", now(), '" + value.Text + "', '" + (comboBox_course.SelectedIndex == 0 ? "201609" : "") + "', '" + comboBox_course.SelectedItem + "')";
            MySqlCommand cmd = new MySqlCommand(query, c);
            cmd.ExecuteNonQuery();
            MessageBox.Show(comboBox_course.SelectedItem + " 과정에 공지사항을 송출했습니다.");
        }
    }

    /*class DBConnection
    {
        public DBConnect()
        {

        }
    }*/ // 추후 DB 연결 관련은 여기서.
}
