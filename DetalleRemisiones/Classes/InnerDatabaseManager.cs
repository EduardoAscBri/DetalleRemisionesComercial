using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DetalleRemisiones.Classes
{
    class InnerDatabaseManager
    {
        public string server;
        public string instance;
        public string database;
        public string user;
        public string password;
        public string connectionString;
        public SqlConnection databaseConnection;
        public InnerDatabaseManager()
        {

        }

        public InnerDatabaseManager(string Servidor, string Instancia, string BaseDeDatos, string Usuario, string Contraseña)
        {
            this.server = Servidor;
            this.instance = Instancia;
            this.database = BaseDeDatos;
            this.user = Usuario;
            this.password = Contraseña;

            createConnectionString();
            createConnectionToDatabase();
        }
        public void createConnectionString()
        {
            this.connectionString = "data source = " + this.server + "\\" + this.instance + "; initial catalog = " + this.database + "; user id = " + this.user + "; password = " + this.password;
        }

        public SqlConnection createConnectionToDatabase()
        {
            this.databaseConnection = new SqlConnection(this.connectionString);

            return this.databaseConnection;
        }

        public SqlConnection createConnectionFromIniFile(string rutaLocal)
        {
            string rutaIni = rutaLocal + "\\" + "Config.ini";
            char splitter = '=';
            string[] valores;
            StreamReader file = new StreamReader(rutaIni);

            valores = file.ReadLine().Split(splitter);
            this.server = valores[1];
            valores = file.ReadLine().Split(splitter);
            this.instance = valores[1];
            valores = file.ReadLine().Split(splitter);
            this.database = valores[1];
            valores = file.ReadLine().Split(splitter);
            this.user = valores[1];
            valores = file.ReadLine().Split(splitter);
            this.password = valores[1];

            createConnectionString();
            createConnectionToDatabase();

            return this.databaseConnection;
        }

        public SqlConnection createConnectionFromIniComercialFile(string rutaLocal)
        {
            string rutaIni = rutaLocal + "\\" + "ConfigComercial.ini";
            char splitter = '=';
            string[] valores;
            StreamReader file = new StreamReader(rutaIni);

            valores = file.ReadLine().Split(splitter);
            this.server = valores[1];
            valores = file.ReadLine().Split(splitter);
            this.instance = valores[1];
            valores = file.ReadLine().Split(splitter);
            this.database = valores[1];
            valores = file.ReadLine().Split(splitter);
            this.user = valores[1];
            valores = file.ReadLine().Split(splitter);
            this.password = valores[1];

            createConnectionString();
            createConnectionToDatabase();

            return this.databaseConnection;
        }

    }
}
