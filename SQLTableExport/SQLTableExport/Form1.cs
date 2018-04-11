using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

using System.Data.SqlClient;

namespace SQLTableExport
{
    public partial class Form1 : Form
    {
        //定義OleDb======================================================
        //1.檔案位置    注意絕對路徑 -> 非 \  是 \\
        private const string FileName = "C:\\Users\\Admin\\Documents\\unium\\MOS工作文件\\摩斯漢堡POS系統檔案結構.xlsx";
        //2.提供者名稱  Microsoft.Jet.OLEDB.4.0適用於2003以前版本，Microsoft.ACE.OLEDB.12.0 適用於2007以後的版本處理 xlsx 檔案
        private const string ProviderName = "Microsoft.ACE.OLEDB.12.0;";
        //3.Excel版本，Excel 8.0 針對Excel2000及以上版本，Excel5.0 針對Excel97。
        private const string ExtendedString = "'Excel 8.0;";
        //4.第一行是否為標題
        private const string Hdr = "No;";
        //5.IMEX=1 通知驅動程序始終將「互混」數據列作為文本讀取
        private const string IMEX = "1';";
        //================================================
        //連線字串
        string cs =
            "Data Source=" + FileName + ";" +
            "Provider=" + ProviderName +
            "Extended Properties=" + ExtendedString +
            "HDR=" + Hdr +
            "IMEX=" + IMEX;
        //Excel 的工作表名稱 (Excel左下角有的分頁名稱)
        string SheetName = "Sheet1";

        public Form1()
        {
            InitializeComponent();
        }

        // 確定按鍵
        private void button1_Click(object sender, EventArgs e)
        {
            OpenSqlConnection();
        }

        // 刪除按鍵
        private void button2_Click(object sender, EventArgs e)
        {
            string queryString = "DELETE FROM TableSchema";
            DeleteCommand(queryString, GetConnectionString());

        }

        // 刪除TableSchema內容
        private static void DeleteCommand(string queryString, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(queryString, connection))
                {
                    try
                    {
                        command.Connection.Open();
                        command.ExecuteNonQuery();
                        MessageBox.Show("刪除完畢");
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        // 更新TableSchema內容
        private void UpdateCommand(SqlConnection connection, DataRow row)
        {
            this.richTextBox1.AppendText("* Update: " + row[0].ToString() + ", " + row[1].ToString() + ", " + row[2].ToString() + ", " + row[3].ToString() + ", " + row[4].ToString() + ", " + row[5].ToString() + ", " + row[6].ToString() + ", " + row[7].ToString() + ", " + row[8].ToString() + "." + System.Environment.NewLine);
            this.richTextBox1.SelectionStart = this.richTextBox1.TextLength;
            this.richTextBox1.ScrollToCaret();

            string queryString = "UPDATE TableSchema SET SeqNo=@SeqNo,_Type=@_Type,_Len=@_Len,_PK=@_PK,_Null=@_Null,_DF=@_DF,_DESC=@_DESC,updDate=@updDate WHERE TableName=@TableName and FieldName=@FieldName";

            using (SqlCommand command = new SqlCommand(queryString, connection))
                {
                    try
                    {
                        command.Parameters.Add("@SeqNo", SqlDbType.Int).Value = row[1].ToString();
                        command.Parameters.Add("@TableName", SqlDbType.VarChar).Value = row[0].ToString();
                        command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = row[2].ToString();
                        command.Parameters.Add("@_Type", SqlDbType.VarChar).Value = row[3].ToString();
                        command.Parameters.Add("@_Len", SqlDbType.Int).Value = row[4].ToString();
                        command.Parameters.Add("@_PK", SqlDbType.Char).Value = row[8].ToString();
                        if (row[5].ToString().Equals("True"))
                        {
                            command.Parameters.Add("@_Null", SqlDbType.Char).Value = 1;
                        }
                        else
                        {
                            command.Parameters.Add("@_Null", SqlDbType.Char).Value = 0;
                        }
                        command.Parameters.Add("@_DF", SqlDbType.Char).Value = row[6].ToString();
                        command.Parameters.Add("@_DESC", SqlDbType.VarChar).Value = row[7].ToString();
                        command.Parameters.Add("@updDate", SqlDbType.DateTime).Value = DateTime.Now.ToString();

                        command.ExecuteNonQuery();
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
        }

        // 新增TableSchema內容
        private void InsertCommand(SqlConnection connection, DataRow row)
        {
            this.richTextBox1.AppendText("+ Insert: " + row[0].ToString() + ", " + row[1].ToString() + ", " + row[2].ToString() + ", " + row[3].ToString() + ", " + row[4].ToString() + ", " + row[5].ToString() + ", " + row[6].ToString() + ", " + row[7].ToString() + ", " + row[8].ToString() + "." + System.Environment.NewLine);
            this.richTextBox1.SelectionStart = this.richTextBox1.TextLength;
            this.richTextBox1.ScrollToCaret();

            string queryString = "INSERT INTO TableSchema(SeqNo,TableName,FieldName,_Type,_Len,_PK,_Null,_DF,_DESC,updDate)" +
                    " VALUES(@SeqNo,@TableName,@FieldName,@_Type,@_Len,@_PK,@_Null,@_DF,@_DESC,@updDate)";

            using (SqlCommand command = new SqlCommand(queryString, connection))
            {
                try
                {
                    command.Parameters.Add("@SeqNo", SqlDbType.Int).Value = row[1].ToString();// "1";
                    command.Parameters.Add("@TableName", SqlDbType.VarChar).Value = row[0].ToString();// "AddCoupon";
                    command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = row[2].ToString();// "PromoNo";
                    command.Parameters.Add("@_Type", SqlDbType.VarChar).Value = row[3].ToString();// "char";
                    command.Parameters.Add("@_Len", SqlDbType.Int).Value = row[4].ToString();// "101";
                    command.Parameters.Add("@_PK", SqlDbType.Char).Value = row[8].ToString();//0;
                    if (row[5].ToString().Equals("True"))
                    {
                        command.Parameters.Add("@_Null", SqlDbType.Char).Value = 1;
                    }
                    else
                    {
                        command.Parameters.Add("@_Null", SqlDbType.Char).Value = 0;
                    }
                    command.Parameters.Add("@_DF", SqlDbType.Char).Value = row[6].ToString();
                    command.Parameters.Add("@_DESC", SqlDbType.VarChar).Value = row[7].ToString();
                    command.Parameters.Add("@updDate", SqlDbType.DateTime).Value = DateTime.Now.ToString();

                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        // 檢查TableSchema內容
        private Boolean SelectCommand(SqlConnection connection, string table, string field)
        {
            /*this.richTextBox1.AppendText("@ SelectCommand: table = " + table + ", field = " + field + "." + System.Environment.NewLine);
            this.richTextBox1.SelectionStart = this.richTextBox1.TextLength;
            this.richTextBox1.ScrollToCaret();*/

            Boolean isFinded = false;
            //using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string queryString = "SELECT * FROM TableSchema WHERE TableName=@TableName and FieldName=@FieldName";

                SqlCommand command = new SqlCommand(queryString, connection);
                command.Parameters.Add("@TableName", SqlDbType.VarChar).Value = table;
                command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = field;

                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    isFinded = true;
                    break;
                }
                reader.Close();
            }
            if (isFinded)
                return true;
            return false;
        }

        // 搜尋所有Table的資料
        private DataSet SelectRows(SqlConnection connection)
        {
            DataSet dataset = new DataSet();
            //using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string queryString = @"SELECT B.[NAME] AS '資料表名稱', " +
                    "C.[COLUMN_ID] AS '欄位順序', " +
                    "C.[NAME] AS '欄位名稱', " +
                    "D.[NAME] AS '資料型別', " +
                    "C.[MAX_LENGTH] AS '長度', " +
                    "C.[IS_NULLABLE] AS '是否允許NULL', " +
                    "E.[TEXT] AS '預設值', " +
                    "( SELECT value " +
                    "FROM fn_listextendedproperty (NULL, 'schema', 'dbo', 'table', B.[NAME], 'column', default) " +
                    "WHERE name='MS_Description' and objtype='COLUMN' and objname Collate Chinese_Taiwan_Stroke_CI_AS=C.[NAME] " +
                    ") AS '欄位備註', " + 
                    "'0' AS 'PK值' " +
                    "FROM SYS.SCHEMAS A INNER JOIN SYS.TABLES B " +
                    "ON A.[SCHEMA_ID] = B.[SCHEMA_ID] " +
                    "INNER JOIN SYS.COLUMNS C " +
                    "ON B.[OBJECT_ID] = C.[OBJECT_ID] " +
                    "INNER JOIN SYS.TYPES D " +
                    "ON C.[SYSTEM_TYPE_ID] = D.[SYSTEM_TYPE_ID] AND C.[USER_TYPE_ID] = D.[USER_TYPE_ID] " +
                    "LEFT JOIN SYS.SYSCOMMENTS E " +
                    "ON C.[DEFAULT_OBJECT_ID] = E.[ID] " +
                    "WHERE B.[TYPE] = 'U' " +
                    "ORDER BY A.[NAME], B.[NAME]";

                SqlDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = new SqlCommand(queryString, connection);
                adapter.Fill(dataset);
            }
            return dataset;
        }

        // 搜尋有PrmaryKey的欄位名稱
        private DataSet SelectPrmaryKeyRaws(SqlConnection connection)
        {
            DataSet dataset = new DataSet();
            string queryString = @"SELECT KU.table_name as TABLENAME,column_name as PRIMARYKEYCOLUMN " +
                "FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS AS TC " +
                "INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS KU " +
                "ON TC.CONSTRAINT_TYPE = 'PRIMARY KEY' AND " +
                "TC.CONSTRAINT_NAME = KU.CONSTRAINT_NAME " +
                "ORDER BY KU.TABLE_NAME, KU.ORDINAL_POSITION";
            
            SqlDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = new SqlCommand(queryString, connection);
            adapter.Fill(dataset);
            return dataset;
        }

        // 開啟SQL連線
        private void OpenSqlConnection()
        {
            string connectionString = GetConnectionString();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    Console.WriteLine("ServerVersion: {0}", connection.ServerVersion);
                    Console.WriteLine("State: {0}", connection.State);

                    Console.Write(DateTime.Now.ToString() + " Read all table data." + "\n");
                    DataSet ds = new DataSet();
                    ds = SelectRows(connection);
                    this.dataGridView1.DataSource = ds.Tables[0].DefaultView;

                    Console.Write(DateTime.Now.ToString() + " Read prmary key data." + "\n");
                    DataSet prmarykeyDataSet = new DataSet();
                    prmarykeyDataSet = SelectPrmaryKeyRaws(connection);
                    foreach(DataRow row in ds.Tables[0].Rows)
                    {
                        foreach(DataRow pkRow in prmarykeyDataSet.Tables[0].Rows)
                        {
                            if (row[0].ToString().Equals(pkRow[0].ToString()) && row[2].ToString().Equals(pkRow[1].ToString()))
                            {
                                row[8] = 1;
                                break;
                            }
                        }
                    }

                    Console.Write(DateTime.Now.ToString() + " Read excel file." + "\n");
                    DataSet fileDataSet = new DataSet();
                    fileDataSet = FileLoad();
                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        foreach (DataRow fileRow in fileDataSet.Tables[0].Rows)
                        {
                            if (row[0].ToString().Equals(fileRow[0].ToString()) && row[2].ToString().Equals(fileRow[1].ToString()))
                            {
                                row[7] = fileRow[2].ToString();
                                break;
                            }
                        }
                    }

                    Console.Write(DateTime.Now.ToString() + " Start update dbo.TableSchema." + "\n");
                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        if (SelectCommand(connection, row[0].ToString(), row[2].ToString()))
                        {
                            UpdateCommand(connection, row);
                        }
                        else
                        {
                            InsertCommand(connection, row);
                        }
                    }
                    Console.Write(DateTime.Now.ToString() + " End update dbo.TableSchema." + "\n");
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private string GetConnectionString()
        {
            // To avoid storing the connection string in your code, 
            // you can retrieve it from a configuration file, using the 
            // System.Configuration.ConfigurationSettings.AppSettings property 
            // 用Windows身分驗證
            //return "Data Source=(local);Initial Catalog=AdventureWorks;" + "Integrated Security=SSPI;";
            // 用SQL Server身分驗證
            //return "Data Source=127.0.0.1;Initial Catalog=mospos;User ID=sa;Password=16308238;";
            return "Data Source=" + this.textBox1.Text + ";Initial Catalog=" + this.textBox4.Text + ";User ID=" + this.textBox2.Text + ";Password=" + this.textBox3.Text + ";";
        }

        // 讀取Excel檔案
        private DataSet FileLoad()
        {
            DataSet dataset = new DataSet();
            using (OleDbConnection cn = new OleDbConnection(cs))
            {
                cn.Open();
                string qs = "select * from[" + SheetName + "$]";
                try
                {
                    using (OleDbDataAdapter dr = new OleDbDataAdapter(qs, cn))
                    {
                        dr.Fill(dataset);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return dataset;
        }


    }
}
