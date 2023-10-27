using Microsoft.Data.Sqlite;
using System.Data;
using System;
using System.IO;
using System.Windows.Forms;
using SQLitePCL;


namespace Conv.Net
{
    internal class DB
    {
        private static SqliteConnection connection = new SqliteConnection("Data Source = dbconvert.db");

        public SqliteConnection GetConnection() => connection;

        public static DataTable ExecuteQuery(string sql, params SqliteParameter[] parameters)
        {
            CreateDatabase();

            SQLitePCL.raw.SetProvider(new SQLitePCL.SQLite3Provider_e_sqlite3());
            SQLitePCL.Batteries.Init();

            var result = new DataTable();
            using (var command = new SqliteCommand(sql, connection))
            {
                foreach (var p in parameters)
                {
                    command.Parameters.Add(p);
                }
                command.Connection.Close();
                command.Connection.Open();
                result.Load(command.ExecuteReader());
                command.Connection.Close();
            }
            return result;
        }
        public static void CloseConnection()
        {
            connection.Dispose();            
        }

        
        public static void CreateDatabase()
        {
            string dbFile = "dbconvert.db";
            if (File.Exists(dbFile) && new FileInfo(dbFile).Length > 0)
            {
                return;
            }

            
            string sqlCommands = @"
                                CREATE TABLE dataconvert (
                                    id INTEGER,
                                    semestr varchar(128),
                                    disciplina varchar(128),
                                    fin varchar(128),
                                    raspred varchar(128),
                                    napravl varchar(128),
                                    kaf	varchar(128),   
                                    groupp varchar(128),
                                    time_all varchar(128),
                                    podgrupp varchar(128),
                                    students varchar(128),
                                    potok varchar(128),
                                    lek varchar(128),
                                    prakt varchar(128),
                                    lab varchar(128),
                                    kursovoi varchar(128),
                                    ucheb_praktika varchar(128),
                                    PRIMARY KEY(id AUTOINCREMENT)                                
                                
                                );";

            
            try
            {
                SqliteConnection conn = new SqliteConnection($"Data Source={dbFile}");
                conn.Open();

                SqliteCommand cmd = new SqliteCommand(sqlCommands, conn);
                cmd.ExecuteNonQuery();

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}. Ошибка создания базы данных. Используйте существующую. ");
            }


        }
    }
}
