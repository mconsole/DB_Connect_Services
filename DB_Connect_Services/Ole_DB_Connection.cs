using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace DB_Connect_Services
{
    public class Ole_DB_Connection
    {
        //OleDb Connection
        public OleDbConnection DbConnection()
        {
            //Microsoft Access Connection String
            string connect = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=..\..\BudgetBuddyTables.accdb;
                             Persist Security Info=False;";
            OleDbConnection con = new OleDbConnection(connect);

            try
            {
                con.Open();
                return con;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return con;
            }
        }


        //Query a database using a SQL string as an input
        public DataTable retrieveDbData(string sqlStr)
        {
            OleDbConnection con1 = DbConnection();
            DataTable tables = new DataTable();

            try
            {
                OleDbCommand cmd = new OleDbCommand(sqlStr, con1);
                cmd.CommandType = CommandType.Text;
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(tables);
                return tables;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return tables;
            }
            finally
            {
                con1.Close();
            }
        }


        //Delete a single database table
        public void deleteTable(string tableNm)
        {
            OleDbConnection con1 = DbConnection();

            try
            {
                string sqlStr = "DELETE * FROM '" + tableNm + "';";
                OleDbCommand cmd = new OleDbCommand(sqlStr, con1);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                con1.Close();
            }

        }


        //Update a single database table
        public void updateTable(int id, string budgetTable, string itemName, string itemType, int itemAmount, string itemUnit, int interest, string itemCategory)
        {
            OleDbConnection con1 = DbConnection();

            try
            {
                string sqlStr = "UPDATE budget_items SET budget_table = '" + budgetTable + "', item_name = '" + itemName + "', item_type = '" + itemType + "', item_amount = " + itemAmount + ", item_unit = '" + itemUnit + "', item_interest = " + interest + ", item_category = '" + itemCategory + "' WHERE ID = " + id + ";";
                OleDbCommand cmd = new OleDbCommand(sqlStr, con1);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                con1.Close();
            }
        }


        //Delete a single database table row
        public void deleteRow(string sqlStr)
        {
            OleDbConnection con1 = DbConnection();

            try
            {
                OleDbCommand cmd = new OleDbCommand(sqlStr, con1);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                con1.Close();
            }
        }
    }
}
