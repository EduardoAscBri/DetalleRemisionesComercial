using ClosedXML.Excel;
using DetalleRemisiones.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DetalleRemisiones.Forms
{
    public partial class Principal : Form
    {
        string rutaLocal;
        SqlConnection conexion;
        SqlCommand command;
        SqlDataAdapter dataAdapter;
        
        public Principal()
        {
            InitializeComponent();
        }

        private void Principal_Load(object sender, EventArgs e)
        {
            this.rutaLocal = Application.StartupPath;
            InnerDatabaseManager databaseManager = new InnerDatabaseManager();
            this.conexion = databaseManager.createConnectionFromIniFile(this.rutaLocal);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string serie = this.txtSerie.Text.Trim();
            string folio = this.txtFolio.Text.Trim();

            //Creacion de la tabla
            DataTable reporte = new DataTable();
            reporte.Columns.Add("Serie");
            reporte.Columns.Add("Folio");
            reporte.Columns.Add("Fecha");
            reporte.Columns.Add("RazonSocial");
            reporte.Columns.Add("RFC");
            reporte.Columns.Add("Cantidad");
            reporte.Columns.Add("Codigo");
            reporte.Columns.Add("Nombre");


            try
            {
                //Productos
                this.command = new SqlCommand("SELECT        admDocumentos.CSERIEDOCUMENTO, admDocumentos.CFOLIO, admDocumentos.CFECHA, " +
                    "admDocumentos.CRAZONSOCIAL, admDocumentos.CRFC, " +
                    "SUM(admMovimientos.CUNIDADES) AS CANTIDAD, Productos.CCODIGOPRODUCTO AS CODIGO, Productos.CNOMBREPRODUCTO AS NOMBRE, Productos.CTIPOPRODUCTO, " +
                    "SUM(admComponentesPaquete.CCANTIDADPRODUCTO * admMovimientos.CUNIDADES) AS CANTIDAD, Componentes.CCODIGOPRODUCTO AS CODIGO, Componentes.CNOMBREPRODUCTO AS NOMBRE " +
                    "FROM            admDocumentos " +
                    "LEFT JOIN admMovimientos " +
                    "ON admDocumentos.CIDDOCUMENTO = admMovimientos.CIDDOCUMENTO " +
                    "LEFT JOIN admProductos AS Productos " +
                    "ON admMovimientos.CIDPRODUCTO = Productos.CIDPRODUCTO " +
                    "LEFT JOIN admComponentesPaquete " +
                    "ON Productos.CIDPRODUCTO = admComponentesPaquete.CIDPAQUETE " +
                    "LEFT JOIN admProductos AS Componentes " +
                    "ON admComponentesPaquete.CIDPRODUCTO = Componentes.CIDPRODUCTO " +
                    "WHERE(admDocumentos.CSERIEDOCUMENTO = @serieDocto AND admDocumentos.CFOLIO = @folioDocto) AND(Productos.CTIPOPRODUCTO != 2) " +
                    "GROUP BY         admDocumentos.CSERIEDOCUMENTO, admDocumentos.CFOLIO, admDocumentos.CFECHA, " +
                    "admDocumentos.CRAZONSOCIAL, admDocumentos.CRFC, " +
                    "Productos.CCODIGOPRODUCTO, Productos.CNOMBREPRODUCTO, Productos.CTIPOPRODUCTO, " +
                    "Componentes.CCODIGOPRODUCTO, Componentes.CNOMBREPRODUCTO", this.conexion);

                this.command.Parameters.AddWithValue("@serieDocto", serie);
                this.command.Parameters.AddWithValue("@folioDocto", folio);
                this.command.CommandTimeout = 600;

                DataTable consultaProductos = new DataTable();
                this.dataAdapter = new SqlDataAdapter(this.command);
                consultaProductos = new DataTable();
                dataAdapter.Fill(consultaProductos);


                //Paquetes
                this.command = new SqlCommand("SELECT        admDocumentos.CSERIEDOCUMENTO, admDocumentos.CFOLIO, admDocumentos.CFECHA, " +
                    "admDocumentos.CRAZONSOCIAL, admDocumentos.CRFC, " +
                    "SUM(admComponentesPaquete.CCANTIDADPRODUCTO * admMovimientos.CUNIDADES) AS CANTIDADCOMPONENTE, Componentes.CCODIGOPRODUCTO AS CODIGOCOMPONENTE, Componentes.CNOMBREPRODUCTO AS NOMBRECOMPONENTE " +
                    "FROM            admDocumentos " +
                    "LEFT JOIN admMovimientos " +
                    "ON admDocumentos.CIDDOCUMENTO = admMovimientos.CIDDOCUMENTO " +
                    "LEFT JOIN admProductos AS Productos " +
                    "ON admMovimientos.CIDPRODUCTO = Productos.CIDPRODUCTO " +
                    "LEFT JOIN admComponentesPaquete " +
                    "ON Productos.CIDPRODUCTO = admComponentesPaquete.CIDPAQUETE " +
                    "LEFT JOIN admProductos AS Componentes " +
                    "ON admComponentesPaquete.CIDPRODUCTO = Componentes.CIDPRODUCTO " +
                    "WHERE(admDocumentos.CSERIEDOCUMENTO = @serieDocto AND admDocumentos.CFOLIO = @folioDocto) AND(Productos.CTIPOPRODUCTO = 2) " +
                    "GROUP BY         admDocumentos.CSERIEDOCUMENTO, admDocumentos.CFOLIO, admDocumentos.CFECHA, " +
                    "admDocumentos.CRAZONSOCIAL, admDocumentos.CRFC, " +
                    "Componentes.CCODIGOPRODUCTO, Componentes.CNOMBREPRODUCTO", this.conexion);

                this.command.Parameters.AddWithValue("@serieDocto", serie);
                this.command.Parameters.AddWithValue("@folioDocto", folio);
                this.command.CommandTimeout = 600;

                DataTable consultaPaquetes = new DataTable();
                this.dataAdapter = new SqlDataAdapter(this.command);
                consultaPaquetes = new DataTable();
                dataAdapter.Fill(consultaPaquetes);

                foreach (DataRow row in consultaProductos.Rows)
                {
                    DataRow dataRow = reporte.NewRow();
                    dataRow["Serie"] = row["CSERIEDOCUMENTO"];
                    dataRow["Folio"] = row["CFOLIO"];
                    dataRow["Fecha"] = row["CFECHA"];
                    dataRow["RazonSocial"] = row["CRAZONSOCIAL"];
                    dataRow["RFC"] = row["CRFC"];
                    dataRow["Cantidad"] = row["CANTIDAD"];
                    dataRow["Codigo"] = row["CODIGO"];
                    dataRow["Nombre"] = row["NOMBRE"];
                    reporte.Rows.Add(dataRow);
                }

                foreach (DataRow row in consultaPaquetes.Rows)
                {
                    DataRow dataRow = reporte.NewRow();
                    dataRow["Serie"] = row["CSERIEDOCUMENTO"];
                    dataRow["Folio"] = row["CFOLIO"];
                    dataRow["Fecha"] = row["CFECHA"];
                    dataRow["RazonSocial"] = row["CRAZONSOCIAL"];
                    dataRow["RFC"] = row["CRFC"];
                    dataRow["Cantidad"] = row["CANTIDADCOMPONENTE"];
                    dataRow["Codigo"] = row["CODIGOCOMPONENTE"];
                    dataRow["Nombre"] = row["NOMBRECOMPONENTE"];
                    reporte.Rows.Add(dataRow);
                }

                XLWorkbook workBook = new XLWorkbook();
                workBook.Worksheets.Add(reporte, "Detalle remision");

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Hoja de calculo Excel (*.xlsx)|*.xlsx";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    workBook.SaveAs(saveFileDialog.FileName);
                    MessageBox.Show("Reporte guardado");
                }
                else
                {
                    MessageBox.Show("Error generando el reporte");
                }
            }
            catch(SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.conexion.Close();
            Application.Exit();
        }
    }
}
