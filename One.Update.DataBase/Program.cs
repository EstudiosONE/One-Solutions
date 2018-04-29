using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace One.Update.DataBase
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Creating DataBase one_iGS... ");
            CreateDataBase();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Done");
            Console.ForegroundColor = ConsoleColor.White;
            Console.ReadLine();

        }
        static void CreateDataBase()
        {
            using (var connection = new SqlConnection("Data Source=DESKTOP-6A8R48U;Initial Catalog=one_iGS;Integrated Security=True"))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = "CREATE DATABASE mydb";
                command.ExecuteNonQuery();
            }
        }
    }
}
