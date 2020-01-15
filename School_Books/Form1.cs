using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

//using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Data.OleDb;
using DataTable = System.Data.DataTable;

namespace School_Books
{
    public partial class Form1 : Form
    {
        readonly string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=books.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=0' ";
        readonly Dictionary<int, string> column =
            new Dictionary<int, string>
            {
                { 1, "Book" },
                { 2, "BorrowDate" },
                { 3, "ReturnDate" }
            };
        public string Sheet { get { return comboBox1.SelectedItem.ToString(); } }
        public string Student { get { return comboBox2.SelectedItem.ToString(); } }

        public Form1()
        {
            InitializeComponent();

        }
        private string[] GetSheets()
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    DataTable dt = new DataTable();
                    dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string[] excelSheets = new String[dt.Rows.Count];
                    int i = 0;
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        excelSheets[i] = dataRow["TABLE_NAME"].ToString().Trim('\'', '$');
                        i++;
                    }
                    return excelSheets;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    return new string[1];

                }
            }
        }
        private string[] GetStudents(string sheet)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string cmd = "SELECT DISTINCT Student FROM [" + sheet + "$]";
                    Debug.WriteLine(cmd);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(cmd, connection);
                    DataTable dt = new DataTable();
                    myDataAdapter.Fill(dt);
                    string[] students = new String[dt.Rows.Count];
                    int i = 0;
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        students[i] = dataRow[0].ToString();
                        i++;
                    }
                    myDataAdapter.Dispose();
                    return students;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    return new string[1];
                }
            }
        }
        private string GetID(string sheet)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string cmd = "SELECT MAX(ID) FROM [" + sheet + "$];";
                    Debug.WriteLine(cmd);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(cmd, connection);
                    DataTable dt = new DataTable();
                    myDataAdapter.Fill(dt);

                    myDataAdapter.Dispose();
                    return dt.Rows[0][0].ToString();

                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    return "";
                }
            }
        }
        private void UpdateGrid(string sheet, string student)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string cmd = "SELECT ID, Book, BorrowDate, ReturnDate FROM [" + sheet + "$] WHERE Student=\"" + student + "\"";
                    Debug.WriteLine(cmd);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(cmd, connection);
                    DataTable dt = new DataTable();
                    myDataAdapter.Fill(dt);
                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns[0].Visible = false;
                    dataGridView1.Columns["Book"].Width = 120;
                    dataGridView1.Columns["Book"].HeaderText = "שם הספר";
                    dataGridView1.Columns["BorrowDate"].Width = 120;
                    dataGridView1.Columns["BorrowDate"].HeaderText = "תאריך השאלה";
                    dataGridView1.Columns["ReturnDate"].Width = 120;
                    dataGridView1.Columns["ReturnDate"].HeaderText = "תאריך החזרה";
                    myDataAdapter.Dispose();
                    return;

                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    return;
                }
            }
        }
        private void UpdateCell(string sheet, DataGridViewCellEventArgs e)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    if (e.RowIndex == -1) { return; }
                    connection.Open();
                    string id = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                    string newValue = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Replace("'", "''");
                    string cmd = "UPDATE [" + sheet + "$] SET " + column[e.ColumnIndex] + " = '" + newValue + "' WHERE ID=" + id;
                    Debug.WriteLine(cmd);
                    OleDbCommand olecmd = new OleDbCommand(cmd, connection);
                    olecmd.ExecuteNonQuery();
                    return;

                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    return;
                }
            }
        }
        private void AddRow(string sheet, string student)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    using (OleDbCommandBuilder bld = new OleDbCommandBuilder(da))
                    {
                        bld.QuotePrefix = "[";  // these are
                        bld.QuoteSuffix = "]";  //   important!

                        da.SelectCommand = new OleDbCommand(
                                "SELECT [ID], [Student], [Book], [BorrowDate], [ReturnDate] " +
                                "FROM [" + sheet + "$] " +
                                "WHERE False",
                                connection);
                        Debug.WriteLine(da.SelectCommand);
                        using (DataTable dt = new System.Data.DataTable("Test"))
                        {
                            // create an empty DataTable with the correct structure

                            da.Fill(dt);
                            DataRow dr = dt.NewRow();

                            dr["ID"] = (Int32.Parse(GetID(sheet)) + 1).ToString();
                            dr["Student"] = student;
                            dr["Book"] = null;
                            dr["BorrowDate"] = null;
                            dr["ReturnDate"] = null;
                            dt.Rows.Add(dr);

                            da.Update(dt);  // write new row back to database
                        }
                    }
                    return;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    return;
                }
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            //ToolTip toolTip1 = new ToolTip
            //{
            //    AutoPopDelay = 5000,
            //    InitialDelay = 1000,
            //    ReshowDelay = 500,
            //    ShowAlways = true
            //};
            //toolTip1.SetToolTip(this.button1, "My button1");

            string[] excelSheets = GetSheets();
            comboBox1.Items.AddRange(excelSheets);
            comboBox1.SelectedIndex = 0;
        }
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Text = "";
            dataGridView1.DataSource = null;
            comboBox2.Items.Clear();
            
            string[] students = GetStudents(Sheet);
            comboBox2.Items.AddRange(students);

        }
        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateGrid(Sheet, Student);
        }
        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            UpdateCell(Sheet, e);
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text.Length == 0)
            {
                return;
            }
            AddRow(Sheet, Student);
            UpdateGrid(Sheet, Student);
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            ComboBox1_SelectedIndexChanged(null, null);
        }
    }
}
