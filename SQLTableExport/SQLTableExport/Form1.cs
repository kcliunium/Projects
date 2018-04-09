using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data.SqlClient;

namespace SQLTableExport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenSqlConnection();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string queryString = "DELETE FROM TableSchema";
            DeleteCommand(queryString, GetConnectionString());
        }

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

        private void UpdateCommand(SqlConnection connection, DataRow row)
        {
            this.richTextBox1.AppendText("* UpdateCommand: " + row[0].ToString() + ", " + row[1].ToString() + ", " + row[2].ToString() + ", " + row[3].ToString() + ", " + row[4].ToString() + ", " + row[5].ToString() + ", " + row[6].ToString() + ", " + row[7].ToString() + "." + System.Environment.NewLine);
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
                        command.Parameters.Add("@_PK", SqlDbType.Char).Value = 1;
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

        private void InsertCommand(SqlConnection connection, DataRow row)
        {
            this.richTextBox1.AppendText("+ InsertCommand: " + row[0].ToString() + ", " + row[1].ToString() + ", " + row[2].ToString() + ", " + row[3].ToString() + ", " + row[4].ToString() + ", " + row[5].ToString() + ", " + row[6].ToString() + ", " + row[7].ToString() + "." + System.Environment.NewLine);
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
                    command.Parameters.Add("@_PK", SqlDbType.Char).Value = 0;
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

        private Boolean SelectCommand(SqlConnection connection, string table, string field)
        {
            this.richTextBox1.AppendText("@ SqlConnection: table = " + table + ", field = " + field + "." + System.Environment.NewLine);
            this.richTextBox1.SelectionStart = this.richTextBox1.TextLength;
            this.richTextBox1.ScrollToCaret();

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

        private DataSet SelectRows(SqlConnection connection, DataSet dataset)
        {
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
                    "B.[CREATE_DATE] AS '建立時間', " +
                    "B.[MODIFY_DATE] AS '修改時間' " +
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
                return dataset;
            }
        }

        private DataSet SelectPrmaryKeyRaws(SqlConnection connection, DataSet dataset)
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
                    "B.[CREATE_DATE] AS '建立時間', " +
                    "B.[MODIFY_DATE] AS '修改時間' " +
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
                return dataset;
        }

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

                    DataSet ds = new DataSet();
                    SelectRows(connection, ds);
                    this.dataGridView1.DataSource = ds.Tables[0].DefaultView;
                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        /*Console.WriteLine(row[0].ToString() + ", " + row[1].ToString() + ", " +
                            row[2].ToString() + ", " + row[3].ToString() + ", " +
                            row[4].ToString() + ", " + row[5].ToString());
                        this.richTextBox1.AppendText("+ InsertData: " + row[0].ToString() + ", " + row[1].ToString() + ", " + row[2].ToString() + ", " + row[5].ToString() + "." + System.Environment.NewLine);
                        this.richTextBox1.SelectionStart = this.richTextBox1.TextLength;
                        this.richTextBox1.ScrollToCaret();*/

                        if (SelectCommand(connection, row[0].ToString(), row[2].ToString()))
                        {
                            UpdateCommand(connection, row);
                        }
                        else
                        {
                            InsertCommand(connection, row);
                        }
                    }
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
    }
}
