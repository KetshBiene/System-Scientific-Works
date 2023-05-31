using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace System_Scientific_Works
{
    public class DB
    {
        private SqlConnection sqlConnection;
        private SqlCommand command;
        private SqlDataReader reader;

        public DB(string connectionString)
        {
            try
            {
                sqlConnection = new SqlConnection(connectionString);
                sqlConnection.Open();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
                reader = null;
            }
        }
        public void BreakCon()
        {
            sqlConnection.Close();
            reader = null;
        }

        public DataSet FullTable(string query)
        {
            if (sqlConnection.State != ConnectionState.Open)
                sqlConnection.Open();

            DataSet ds = new DataSet();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(query, sqlConnection);

            dataAdapter.Fill(ds);

            return ds;
        }

        public void DoQuery(string query) 
        {
            if (sqlConnection.State != ConnectionState.Open)
                sqlConnection.Open();

            command = new SqlCommand(query, sqlConnection);
            command.ExecuteNonQuery();
        }

        public int GetLastId(string from, string idname)
        {

            if (sqlConnection.State != ConnectionState.Open)
                sqlConnection.Open();

            command = new SqlCommand($"WITH SRC AS (SELECT TOP(1) {idname} FROM {from} ORDER BY {idname} DESC) SELECT * FROM SRC", sqlConnection);
            reader = command.ExecuteReader();

            int a = 0; 
            while (reader.Read())
            {
                a = reader.GetInt32(0);
            }
                
            reader.Close();
            return a;           
        }

        public Data GetNameId(string tableName)
        {
            List<string> names = new List<string>();
            List<int> ids = new List<int>();

            if (sqlConnection.State != ConnectionState.Open)
                sqlConnection.Open();
            
            command = new SqlCommand($"SELECT Name, Id FROM {tableName}", sqlConnection);
         
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                names.Add(reader.GetString(0));
                ids.Add(reader.GetInt32(1));
            }

            reader.Close();
            return new Data(names, ids);
        }
        public Data GetNameId(string tableName, string columns)
        {
            List<string> names = new List<string>();
            List<int> ids = new List<int>();

            if (sqlConnection.State != ConnectionState.Open)
                sqlConnection.Open();

            command = new SqlCommand($"SELECT {columns} FROM {tableName}", sqlConnection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                names.Add(reader.GetString(0));
                ids.Add(reader.GetInt32(1));
            }

            reader.Close();
            return new Data(names, ids);
        }
        public Data GetNameId(string tableName, int id)
        {
            List<string> names = new List<string>();
            List<int> ids = new List<int>();

            if (sqlConnection.State != ConnectionState.Open)
                sqlConnection.Open();

            command = new SqlCommand($"SELECT Name, Id FROM {tableName} WHERE Faculty_Id={id}", sqlConnection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                names.Add(reader.GetString(0));
                ids.Add(reader.GetInt32(1));
            }

            reader.Close();
            return new Data(names, ids);
        }

        public void Update(string tablename, string what, string where)
        {
            if (sqlConnection.State != ConnectionState.Open)
                sqlConnection.Open();

            command = new SqlCommand($"UPDATE {tablename} SET {what} WHERE {where}", sqlConnection);
            command.ExecuteNonQuery();
        }

        public void Delete(string tablename, string where)
        {
            if (sqlConnection.State != ConnectionState.Open)
                sqlConnection.Open();

            command = new SqlCommand($"DELETE FROM {tablename} WHERE {where}", sqlConnection);
            command.ExecuteNonQuery();
        }

    }
}
