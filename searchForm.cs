using System;
using System.Data;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace adminPrg
{ 
    public partial class searchForm : Form
    {

        DBConnect db = new DBConnect();
        ReserveData rd = new ReserveData();
        TextBox[] tx; 
        bool ischecked_delete = false;
        int dataSize = -1, changedDataCount = 0, oldCount = 0;
        public int[] changedIndex_column, changedIndex_row;

        public searchForm()
        {
            InitializeComponent();
            portnumber.Text = "";
            servernumber.Text = "";
            data_id.Text = "";
            data_pw.Text = "";
            tx = new TextBox[] { insert_studid, insert_deptname, insert_name, phone, insert_grade };
        }

        public DataGridView selectedView()
        {
            if (tabControl.SelectedTab == tabPage1)
                return showdata_1;
            else
                return showdata_2;
        }

        private string makeCondition()
        {
            string s = "";

            if (deptname.Text != "")
            {
                s += "student_dept = '" + deptname.Text + "'";
                if (studid.Text != "" || name.Text != "" || phone.Text != "" )
                    s += " AND ";
            }

            if (studid.Text != "")
            {
                s += "student_num = " + studid.Text;
                if (name.Text != "" || phone.Text != "")
                    s += " AND ";
            }

            if (name.Text != "")
            {
                s += "student_name = '" + name.Text + "'";
                if (phone.Text != "")
                    s += " AND ";
            }

            if (phone.Text != "")
                s += "student_phone = " + phone.Text;

            return s;
        }

        public void Initialize_Array(int dataSize = 0)
        {
            changedIndex_column = null;
            changedIndex_column = new int[dataSize];
            init_Array(ref changedIndex_column);

            changedIndex_row = null;
            changedIndex_row = new int[dataSize];
            init_Array(ref changedIndex_row);
        }

        private void init_Array(ref int[] arr)
        {
            for (int i = 0; i < dataSize; i++)
                arr[i] = -1;
        }

        private void read_db_Click(object sender, EventArgs e)
        {
            if (servernumber.Text == "" || portnumber.Text == "" || data_id.Text == "" || data_pw.Text == "")
                MessageBox.Show("접속 정보가 입력되어야 합니다.");
            else {
                db.Connect(connection_status, servernumber.Text, portnumber.Text, data_id.Text, data_pw.Text, dataGridView_timetable);
                combo_course.Enabled = true;
                savetoserver.Enabled = true;
                rd.getDB(db, selectedView(), ref dataSize, ref oldCount, combo_loadTable.SelectedIndex);
                Initialize_Array(dataSize);
            }
        }

        private void savetoserver_Click(object sender, EventArgs e)
        {
            db.Update(changedIndex_column, changedIndex_row, oldCount, combo_loadTable.SelectedIndex ,showdata_1);
            savetoserver.Enabled = false;
            changedDataCount = 0;
            init_Array(ref changedIndex_column);
            init_Array(ref changedIndex_row);
            show_dataChanged.Text = 0.ToString() + " Cells";
            show_target.Items.Clear();
        }
        
        private void openexcel_Click(object sender, EventArgs e)
        {
            rd.setExcelData(selectedView(), selectedView().ColumnCount);
        }

        private void save_Click(object sender, EventArgs e)
        {
            rd.save_to_Excelfile(selectedView());
        }

        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            db.viewType = (selectedView() == showdata_1 ? true : false);
        }

        private void closedb_Click(object sender, EventArgs e)
        {
            //db.Disconnect(connection_status);
        }

        private void searchtocondition_Click(object sender, EventArgs e)
        {
            rd.getDB(db, selectedView(), ref dataSize, ref oldCount, combo_loadTable.SelectedIndex ,makeCondition());
        }

        private void insert_button_Click(object sender, EventArgs e)
        {
            if (!ischecked_delete)
                db.Insert(tx);
            else
                db.Delete(tx, makeCondition());
        }

        private void checked_delete_CheckedChanged(object sender, EventArgs e)
        {
            if (ischecked_delete)
            {
                ischecked_delete = false;
                insert_button.Text = "입력";
            }
            else
            {
                ischecked_delete = true;
                insert_button.Text = "삭제";
            }
        }

        private void showdata_1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            changedIndex_column[changedDataCount] = showdata_1.CurrentCell.ColumnIndex;
            changedIndex_row[changedDataCount] = showdata_1.CurrentCell.RowIndex;

            int tmp = -1;
            int tmp2 = -1; // 기준값

            if (changedDataCount < dataSize)
            {
                for (int i = 0; i < changedDataCount; i++)
                {
                    tmp = changedIndex_row[i];
                    tmp2 = changedIndex_column[i];

                    for (int j = i + 1; j < changedDataCount; j++)
                    {
                        if (tmp == changedIndex_row[j] && tmp2 == changedIndex_column[j])
                        {
                            for (int k = j; k < changedDataCount; k++)
                            {
                                changedIndex_row[k] = changedIndex_row[k +1];
                                changedIndex_column[k] = changedIndex_column[k +1];
                            }
                            changedDataCount--;
                            break;
                        }
                    }
                } // 중복 제거
                show_target.Items.Add(changedIndex_column[changedDataCount] + ", " + changedIndex_row[changedDataCount]);
                changedDataCount++;
            }
            show_dataChanged.Text = changedDataCount + " Cells";
        }

        private void combo_course_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(combo_course.SelectedIndex) {
                case 0:
                    dataGridView_timetable.DataSource = null;
                    break;
                default:
                    db.changedCourse(dataGridView_timetable, combo_course.SelectedItem.ToString());
                    break;
            }
        }

        private void insert_score_TextChanged(object sender, EventArgs e)
        {

        } 
    }

    class DBConnect
    {
        public bool viewType = true;
        public int openType = -1;
        public delegate string setter(string data);

        MySqlConnection c;

        public DBConnect()
        {
            
        }

        public void tableType(bool type)
        {
            this.viewType = type;
        }

        public void Connect(System.Windows.Forms.TextBox tx, string server, string port, string id, string pw, DataGridView g)
        {
            String connector = "Server=" + server + "; Port=" + port + "; database=mysql; uid=" + id + "; pwd=" + pw + "; charset=utf8;";

            try
            {
                c = new MySqlConnection(connector);
                c.Open();
                tx.Clear();
                tx.AppendText("Connected");                
            }
            catch (MySqlException ex)
            {
                tx.Clear();
                tx.AppendText("Error : " + ex.Message);
                return;
            }
        }

        public void changedCourse(DataGridView g, string cName)
        {
            DataSet ds = new DataSet();
            string day = DateTime.Now.DayOfWeek.ToString();
            int day_int = -1;

            switch(day)
            {
                case "Monday":
                    day_int = 1; break;
                case "Tuesday":
                    day_int = 2; break;
                case "Wednesday":
                    day_int = 3; break;
                case "Thursday":
                    day_int = 4; break;
                case "Friday":
                    day_int = 5; break;
            }

            MySqlDataAdapter ad = new MySqlDataAdapter(@"SELECT day, start_time, end_time, cl_code FROM COURSE_TIMETABLE where course = '" + cName +"' AND day = " + day_int + ";", c);
            ad.Fill(ds);

            g.DataSource = null;

            for (int i=0; i<ds.Tables[0].Columns.Count; i++)
                ds.Tables[0].Columns[i].ColumnName = setTableName(ds.Tables[0].Columns[i].ColumnName); // 컬럼명 설정
            
            for(int i=0; i<ds.Tables[0].Rows.Count; i++)
                    ds.Tables[0].Rows[i][0] = setDay(ds.Tables[0].Rows[i][0].ToString()); // 요일 설정

            g.DataSource = ds.Tables[0];
            g.ReadOnly = true;
            g.Columns[0].Width = 55;
            g.Columns[1].Width = 80;
            g.Columns[2].Width = 80;
            g.Columns[3].Width = 69;
        }

        /*public void Disconnect(System.Windows.Forms.TextBox tx)
        {
            c.Close();
            tx.ResetText();
            tx.AppendText("Disconnect");
        }*/

        private string setTableName(string name)
        {
            switch(name)
            {
                case "student_num":
                    return "학번";
                case "student_name":
                    return "이름";
                case "student_grade":
                    return "학년";
                case "student_dept":
                    return "학과";
                case "student_phone":
                    return "전화번호";
                case "start_time":
                    return "시작시간";
                case "end_time":
                    return "종료시간";
                case "cl_code":
                    return "강의실";
                case "day":
                    return "요일";
                case "id_num":
                    return "작성번호";
                case "time_stamp":
                    return "작성일";
                case "notice_char":
                    return "공지 내용";
                case "se_code":
                    return "학기코드";
                case "course":
                    return "과정";
                case "st_course":
                    return "학생과정";
                case "st_course_id":
                    return "학생과정";
                case "point":
                    return "취득점수";
                case "session":
                    return "시험회차";
                case "test_type":
                    return "시험종류";
            }
            return "";
        }

        private string setDay(string day) {
            switch(day)
            {
                case "1":
                    return "월";
                case "2":
                    return "화";
                case "3":
                    return "수";
                case "4":
                    return "목";
                case "5":
                    return "금";
            }
            return "";
        }

        public DataSet Select(DataSet ds, int o_type, string s="")
        {
            string query = @"SELECT ";

            if(viewType)
            {
                switch (o_type) {
                    case 0:
                        query += "* FROM STUDENT_INFO" + (viewType == true && s != "" ? " WHERE " + s + "; " : "; ");
                        break;
                    case 1:
                        query += "* FROM STUDENT_POINT;";
                        break;
                    case 2:
                        query += "* FROM STUDENT_COURSE;";
                        break;
                }
            } else
                query += "id_num, time_stamp, notice_char FROM NOTICE;";

            MySqlDataAdapter adapter = new MySqlDataAdapter(query, c);
            adapter.Fill(ds);

            for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                ds.Tables[0].Columns[i].ColumnName = setTableName(ds.Tables[0].Columns[i].ColumnName);

            return ds;
        }

        public void Insert(TextBox[] tx)
        {
            if(tx[0].Text == "" || tx[1].Text == "" || tx[2].Text == "" || tx[3].Text == "" || tx[4].Text == "")
            {
                MessageBox.Show("모든 항목을 입력하셔야 합니다.", "Information");
                return;
            }

            string query = "INSERT INTO STUDENT_INFO VALUES(" + Int32.Parse(tx[0].Text) +", '" + tx[2].Text + "', " + Int32.Parse(tx[4].Text) +", '" + tx[1].Text + "', '" + (tx[3].Text == "" ? "null" : tx[3].Text) + "');";
            MySqlCommand cmd = new MySqlCommand(query, c);
            cmd.ExecuteNonQuery();
            MessageBox.Show(tx[2].Text + "학생의 정보가 입력되었습니다.");
        }

        public void Delete(TextBox[] tx, string where)
        {
           string query = "DELETE FROM " + (viewType ? "STUDENT_INFO" : "NOTICE") + (where != "" ? "WHERE " + where + ";" : ";");
           MySqlCommand cmd = new MySqlCommand(query, c);
           cmd.ExecuteNonQuery();
           
           MessageBox.Show(tx[3].Text + "학생의 정보가 삭제되었습니다.");
        }

        public void Update(int[] col, int[] row, int size, int v_type, DataGridView g)
        {
            MySqlCommand cmd = null;

            int updateCount = 0;

            for (int i = 0; i < row.Length; i++)
            {
                if (row[i] == -1)
                    break;
                else
                    updateCount++;
            } 

            string query = "";

            for (int i = 0; i < updateCount; i++)
            {
                try
                {

                    switch (v_type)
                    {
                        case 0:
                            if (g.Rows[row[i]].Cells[1].Value.ToString() == "" && g.Rows[row[i]].Cells[2].Value.ToString() == ""
                                                                               && g.Rows[row[i]].Cells[3].Value.ToString() == ""
                                                                               && g.Rows[row[i]].Cells[4].Value.ToString() == "")
                                query = "DELETE FROM STUDENT_INFO where student_num = " + g.Rows[row[i]].Cells[0].Value + ";";
                            else
                            {
                                query = "UPDATE STUDENT_INFO SET student_num = " + g.Rows[row[i]].Cells[0].Value
                                                 + ", student_name = '" + g.Rows[row[i]].Cells[1].Value
                                                 + "', student_grade = " + g.Rows[row[i]].Cells[2].Value
                                                 + ", student_dept = '" + g.Rows[row[i]].Cells[3].Value
                                                 + "', student_phone = '" + g.Rows[i].Cells[4].Value
                                                 + "' where student_num = " + g.Rows[row[i]].Cells[0].Value + ";";
                            }
                            break;
                        case 1:
                            if (g.Rows[row[i]].Cells[1].Value.ToString() == "" && g.Rows[row[i]].Cells[2].Value.ToString() == ""
                                                                               && g.Rows[row[i]].Cells[3].Value.ToString() == "")
                                query = "DELETE FROM STUDENT_POINT where st_course_id = '" + g.Rows[row[i]].Cells[0].Value + "';";
                            else
                            {
                                query = "UPDATE STUDENT_POINT SET st_course = '" + g.Rows[row[i]].Cells[0].Value
                                                 + "', point = " + g.Rows[row[i]].Cells[1].Value
                                                 + ", session = '" + g.Rows[row[i]].Cells[2].Value
                                                 + "', test_type = '" + g.Rows[row[i]].Cells[3].Value
                                                 + "' where st_course_id = '" + g.Rows[row[i]].Cells[0].Value + "' AND test_type = '" + g.Rows[row[i]].Cells[3].Value + "';";
                            }
                            break;
                        case 2:
                            if (g.Rows[row[i]].Cells[1].Value.ToString() == "" && g.Rows[row[i]].Cells[2].Value.ToString() == ""
                                                                               && g.Rows[row[i]].Cells[3].Value.ToString() == "")
                                query = "DELETE FROM STUDENT_COURSE where student_num = '" + g.Rows[row[i]].Cells[0].Value + "';";
                            else
                            {
                                query = "UPDATE STUDENT_COURSE SET student_num = '" + g.Rows[row[i]].Cells[0].Value
                                                 + "', course = '" + g.Rows[row[i]].Cells[1].Value
                                                 + "', se_code = '" + g.Rows[row[i]].Cells[2].Value
                                                 + "', st_course = '" + g.Rows[row[i]].Cells[3].Value
                                                 + "' where student_num = '" + g.Rows[row[i]].Cells[0].Value + "';";
                            }
                            break;
                        case 3:
                            break;
                        case 4:
                            break;
                    }
                } catch (Exception e) { }
                
                cmd = new MySqlCommand(query, c);
                cmd.ExecuteNonQuery();
            }

            if (size <= g.Rows.Count - 1)
            {
                for (int i = size; i < g.Rows.Count - 1; i++)
                {
                    //MessageBox.Show("Insertion - " + i + "번째 위치");

                    switch(v_type)
                    {
                        case 0:
                            query = "INSERT INTO STUDENT_INFO VALUES(" + g.Rows[i].Cells[0].Value
                                                                       + ", '" + g.Rows[i].Cells[1].Value
                                                                       + "', " + g.Rows[i].Cells[2].Value
                                                                       + ", '" + g.Rows[i].Cells[3].Value
                                                                       + "', '" + g.Rows[i].Cells[4].Value + "');";
                            break;
                        case 1:
                            query = "INSERT INTO STUDENT_POINT VALUES('" + g.Rows[i].Cells[0].Value
                                                                         + "', " + g.Rows[i].Cells[1].Value
                                                                         + ", '" + g.Rows[i].Cells[2].Value
                                                                         + "', '" + g.Rows[i].Cells[3].Value + "');";
                            break;
                        case 2:
                            query = "INSERT INTO STUDENT_COURSE VALUES('" + g.Rows[i].Cells[0].Value
                                                                         + "', '" + g.Rows[i].Cells[1].Value
                                                                         + "', '" + g.Rows[i].Cells[2].Value
                                                                         + "', '" + g.Rows[i].Cells[3].Value + "');";
                            break;
                }
                    cmd = new MySqlCommand(query, c);
                    cmd.ExecuteNonQuery();
                }
            }// 새로 추가된 열이 있으므로 insert문만 접근
        }
    }

    partial class ReserveData 
    {
        public bool opentype = false;

        public void getDB(DBConnect db, DataGridView g, ref int arrSize, ref int old, int o_type, string s = "")
        {
            DataSet ds = new DataSet();

            opentype = true;
           
            g.Columns.Clear();
            g.DataSource = null; // 뭘로 지우면 좋을 지 몰라서 둘 다 씀.

            db.Select(ds, o_type, s);

            old = ds.Tables[0].Rows.Count;

            if(arrSize == -1)
                arrSize = (ds.Tables[0].Rows.Count + 1) * (ds.Tables[0].Columns.Count + 1);

            g.DataSource = ds.Tables[0];

            g.ReadOnly = (db.viewType ? false : true);

        }
}
