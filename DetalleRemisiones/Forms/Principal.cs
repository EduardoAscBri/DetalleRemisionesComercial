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
        DataTable dataTable = new DataTable();
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
            this.dataTable.Clear();

            string serie = this.txtSerie.Text.Trim();
            string folio = this.txtFolio.Text.Trim();


            try
            {
                this.command = new SqlCommand("SELECT        admDocumentos.CSERIEDOCUMENTO, admDocumentos.CFOLIO, admDocumentos.CFECHA, " +
                    "admDocumentos.CRAZONSOCIAL, admDocumentos.CRFC, " +
                    "SUM(admComponentesPaquete.CCANTIDADPRODUCTO * admMovimientos.CUNIDADES) AS CANTIDADCOMPONENTES, Componentes.CCODIGOPRODUCTO AS CODIGOCOMPONENTE, Componentes.CNOMBREPRODUCTO AS NOMBRECOMPONENTE " +
                    "FROM            admDocumentos " +
                    "LEFT JOIN admMovimientos " +
                    "ON admDocumentos.CIDDOCUMENTO = admMovimientos.CIDDOCUMENTO " +
                    "LEFT JOIN admProductos AS Paquetes " +
                    "ON admMovimientos.CIDPRODUCTO = Paquetes.CIDPRODUCTO " +
                    "LEFT JOIN admComponentesPaquete " +
                    "ON Paquetes.CIDPRODUCTO = admComponentesPaquete.CIDPAQUETE " +
                    "LEFT JOIN admProductos AS Componentes " +
                    "ON admComponentesPaquete.CIDPRODUCTO = Componentes.CIDPRODUCTO " +
                    "WHERE(admDocumentos.CSERIEDOCUMENTO = @serieDocto AND admDocumentos.CFOLIO = @folioDocto) " +
                    "GROUP BY         admDocumentos.CSERIEDOCUMENTO, admDocumentos.CFOLIO, admDocumentos.CFECHA, " +
                    "admDocumentos.CRAZONSOCIAL, admDocumentos.CRFC, " +
                    "Componentes.CCODIGOPRODUCTO, Componentes.CNOMBREPRODUCTO", this.conexion);

                this.command.Parameters.AddWithValue("@serieDocto", serie);
                this.command.Parameters.AddWithValue("@folioDocto", folio);
                this.command.CommandTimeout = 600;

                this.dataAdapter = new SqlDataAdapter(this.command);
                this.dataTable = new DataTable();
                this.dataAdapter.Fill(dataTable);


                XLWorkbook workBook = new XLWorkbook();
                workBook.Worksheets.Add(dataTable, "Detalle remision");

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
