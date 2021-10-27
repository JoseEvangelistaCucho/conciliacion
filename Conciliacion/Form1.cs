using Conciliacion.DATA;
using Conciliacion.Model;
using ExcelDataReader;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Conciliacion
{

    public partial class Form1 : Form
    {
        ModelConexion mc = new ModelConexion();
        private DataSet dt = new DataSet();
        public string ruta;
        public Form1()
        {
            InitializeComponent();
            btnConciliar.Enabled = false;
            btnMostrar.Enabled = false;
            btnGuardar.Enabled = false;
            txtRuta.Enabled = false;
            //   txtUrl.Enabled = false;
           // btnBantotal.Enabled = false;
            dateTimePicker1.Enabled = false;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            btnConciliar.Cursor = Cursors.Hand;
            btnImportar.Cursor = Cursors.Hand;
            txtRuta.Enabled = true;
            dgvDatos.SelectionMode = DataGridViewSelectionMode.
            FullRowSelect;
            dgvDatos.MultiSelect = false;
            dgvDatos.ReadOnly = true;
            dgvDatos.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvDatos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvDatos.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
        }

        private void btnImportar_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Excel Workbook|*.xlsx";
                dateTimePicker1.Enabled = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    cboHojas.Items.Clear();
                    dgvDatos.DataSource = null;
                    txtRuta.Text = openFileDialog1.FileName;
                    FileStream fsSource = new FileStream(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader = ExcelReaderFactory.CreateReader(fsSource);
                    ruta = Path.GetFileName(txtRuta.Text);
                    var r = ruta.Split('.');
                    txtRuta.Text = r[0].ToString();

                    dt = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    foreach (DataTable tabla in dt.Tables)
                    {
                        cboHojas.Items.Add(tabla.TableName);
                    }
                    cboHojas.SelectedIndex = 0;

                    reader.Close();
                }
                if (ruta != null)
                {

                    btnMostrar.Enabled = true;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Verifique que el Archivo este Cerrado");
                txtRuta.Text = "";
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            
            ////////
              string selectDate = dateTimePicker1.Value.ToString("MM/dd/yyyy");
             var c = "Data Source=10.105.0.4;Initial Catalog=Bantotal;Persist Security Info=True;User ID=DSEVOL;Password=BancoDigital2809";
              SqlConnection co = new SqlConnection(c);

          //  SqlConnection co = new SqlConnection(Conexion.cnn);
            //////
            ///// Data
            ////
            dgvDatos.Columns.Clear();
            string comando = "select jazc34idre as 'Id Remesa', jazc34tip as 'Tipo (Env/Rec)'," +
                " CONVERT(varchar,jazc34fcc,103)  as 'Fecha',jazc34idER as 'Empresa',  jazc34mdp as 'Mda Pago', jazc34imp as 'Importe', jazc34ref as MTCN, " +
                "jazc34est as 'Estado', REPLACE(JAZC34NDoO,' ','') as 'NDoc Ordenante', REPLACE(JAZC34NDoB,' ','') as 'NDoc Beneficiario'   from JAZC34 where jazc34fcc=@fecha and jazc34est='C'";
            SqlCommand cmd = new SqlCommand(comando, co);
            cmd.Parameters.AddWithValue("@fecha", selectDate);

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dat = new DataTable();
            da.Fill(dat);

            da.Dispose();
            co.Close();

            dgvBanTotal.DataSource = dat;
            
            

            ////////
            string selectDateAsString = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string cx = mc.url;
            dateTimePicker1.Enabled = true;


            if (cboHojas.Text=="")
            {
                MessageBox.Show("Seleccione una hoja");
                return;
            }
           
            if (cx != null)
            {
                Conexion.cnn = cx;
            }
            SqlConnection cn = new SqlConnection(Conexion.cnn);
            try
            {
                btnMostrar.Enabled = false;
                btnMostrar.Text = "Mostar";
                //
                DataTable dt = (DataTable)(dgvBanTotal.DataSource);
                bool res = new Operaciones().cargarTabla(dt);
                // var r = ruta.Split('.');
                if (cboHojas.Text == "Sheet1")
                {
                    DataTable data = (DataTable)(dgvDatos.DataSource);
                    bool resultado = new Operaciones().cargarDatos(data);

                    if (resultado)
                    {
                        /////
                        //// NO CONCILIARON
                        ///
                        dgvDatos.Columns.Clear();
                   
                            SqlCommand comandoRia = new SqlCommand("SELECT  t1.*,t2.jazc34idre as 'Id Remesa BT'," +
                            " t2.jazc34tip as 'Tipo (Env/Rec) BT', t2.jazc34fcc as 'Fecha BT',t2.jazc34mdp as 'Mda Pago BT'," +
                            " t2.jazc34imp as 'Importe BT',t2.jazc34ref as 'MTCN BT', t2.jazc34est as 'Estado BT'," +
                            " t2.JAZC34NDoO as 'NDoc Ordenante BT', t2.JAZC34NDoB as 'NDoc Beneficiario BT',t2.jazc34idER as 'EMPRESA BT' " +
                                "FROM conciliacion T1 FULL OUTER JOIN JAZC34 T2 " +
                                            "   ON  T1.pin = T2.jazc34ref WHERE "+
             " ISNULL(CONVERT(varchar(12), CAST(T1.[importe_pago] AS decimal(20, 2))), -1) != "+
              "  ISNULL(CONVERT(varchar(12), CAST(T2.[jazc34imp] AS decimal(20, 2))), -1) or  T1.pin != T2.jazc34ref", cn);
                        // SqlCommand cmd = new SqlCommand(comandoRia, cn);
                        comandoRia.Parameters.AddWithValue("@fecha", selectDateAsString);
                        SqlDataAdapter adaptador = new SqlDataAdapter(comandoRia);
                      //  adaptador.SelectCommand = comandoRia;
                        DataTable tabla = new DataTable();
                        adaptador.Fill(tabla);
                        dgvDatos.DataSource = tabla;

                        //////
                        ///// CONCILIARION
                        ////
                        dgvIguales.Columns.Clear();
                            SqlCommand comandoIguales = new SqlCommand("SELECT  CAST(t1.fecha AS varchar(10)) as 'Fecha', 'none' as 'Remitente', t1.beneficiario as 'Beneficiario', " +
                        " t1.pin as 'MTCN', t1.seq as 'Seq', t1.importe_pago as 'Principal', " +
                        " t1.comision as 'Cargo', t2.jazc34idER as 'EMPRESA',t1.importe_pago as 'Total Enviado', GETDATE()as 'Fecha Consulta' " +
                        "  FROM conciliacion T1 INNER JOIN JAZC34 T2 " +
                        "   ON  T1.pin = T2.jazc34ref WHERE   t2.jazc34est = 'C' " +
                        " and ISNULL(CONVERT(varchar(12), CAST(T1.[importe_pago] AS decimal(20, 2))), -1) = " +
                        "  ISNULL(CONVERT(varchar(12), CAST(T2.[jazc34imp] AS decimal(20, 2))), -1) " + " drop table conciliacion", cn);
                        //  SqlCommand cmd2 = new SqlCommand(comandoRia, cn);
                        comandoIguales.Parameters.AddWithValue("@fecha", selectDateAsString);
                        SqlDataAdapter a = new SqlDataAdapter(comandoIguales);
                     //   a.SelectCommand = comandoIguales;
                        DataTable tabla2 = new DataTable();
                        a.Fill(tabla2);
                        dgvIguales.DataSource = tabla2;
                        ////////
                        ///
                       SqlCommand eliminartabla = new SqlCommand("drop table conciliacion", cn);
                        SqlDataAdapter ad = new SqlDataAdapter();
                        ad.SelectCommand = eliminartabla;                      

                        MessageBox.Show("Se Concilio con Exito");
                        btnGuardar.Enabled = true;
                        btnConciliar.Enabled = false;
                    }

                    else
                    {
                        MessageBox.Show("Hubo un problema al Conciliar");
                    }
                }
                if (cboHojas.Text == "Maestro OutBound Dolares")
                {
                    DataTable data = (DataTable)(dgvDatos.DataSource);

                    bool resultado = new Operaciones().cargarDatosWu(data);

                    if (resultado)
                    {
                        /////
                        //// NO CONCILIARON
                        ///
                        dgvDatos.Columns.Clear();
                        //   SqlConnection cn = new SqlConnection(Conexion.cnn);
                        SqlCommand comandoRia = new SqlCommand(" SELECT   t1.*,t2.jazc34idre as 'Id Remesa BT'," +
                            " t2.jazc34tip as 'Tipo (Env/Rec) BT', t2.jazc34fcc as 'Fecha BT',t2.jazc34mdp as 'Mda Pago BT'," +
                            " t2.jazc34imp as 'Importe BT',t2.jazc34ref as 'MTCN BT', t2.jazc34est as 'Estado BT'," +
                            " t2.JAZC34NDoO as 'NDoc Ordenante BT', t2.JAZC34NDoB as 'NDoc Beneficiario BT',t2.jazc34idER as 'EMPRESA BT' " +
                            " FROM conciliacionwu T1 FULL OUTER JOIN JAZC34 T2 " +
                            " ON  T1.MTCN = T2.jazc34ref  WHERE   " +
                            " ISNULL(CONVERT(varchar(12), CAST(T1.[Principal] AS decimal(20, 2))), -1) != " +
                            " ISNULL(CONVERT(varchar(12), CAST(T2.[jazc34imp] AS decimal(20, 2))), -1) or  T1.MTCN != T2.jazc34ref ", cn);
                        comandoRia.Parameters.AddWithValue("@fecha", selectDateAsString);
                        SqlDataAdapter adaptador = new SqlDataAdapter(comandoRia);
                      //  adaptador.SelectCommand = comandoRia;
                        DataTable tabla = new DataTable();
                        adaptador.Fill(tabla);
                        dgvDatos.DataSource = tabla;

                        //////
                        ///// CONCILIARION
                        ////
                        dgvIguales.Columns.Clear();
                        SqlCommand comandoIguales = new SqlCommand(" SELECT CAST(t1.fecha AS varchar(10))  as 'Fecha', concat(t1.apellido_Beneficiario, ' ', t1.Nombre_Beneficiario) as 'Beneficiario', " +
                            " t1.Remitente as 'Remitente', t1.MTCN as 'MTCN', t1.Principal as 'Principal', t1.Cargo as 'Cargo', t1.IGV, t1.ITF, " +
                            " t1.Total_Enviado as 'Total Enviado', t1.Pais_Beneficiario , t2.jazc34idER as 'EMPRESA' ,GETDATE()as 'Fecha Consulta' " +
                            "   FROM conciliacionwu T1 INNER JOIN JAZC34 T2 " +
                            " ON  T1.MTCN = T2.jazc34ref  WHERE " +
                            " ISNULL(CONVERT(varchar(12), CAST(T1.[Principal] AS decimal(20, 2))), -1) = " +
                            " ISNULL(CONVERT(varchar(12), CAST(T2.[jazc34imp] AS decimal(20, 2))), -1) " + " Drop table conciliacionwu", cn);
                        //  SqlCommand cmd2 = new SqlCommand(comandoRia, cn);
                        comandoIguales.Parameters.AddWithValue("@fecha", selectDateAsString);
                        SqlDataAdapter a = new SqlDataAdapter(comandoIguales);
                        //   a.SelectCommand = comandoIguales;
                        DataTable tabla2 = new DataTable();
                        a.Fill(tabla2);
                        dgvIguales.DataSource = tabla2;


                        MessageBox.Show("Se Concilio con Exito");
                        btnGuardar.Enabled = true;
                        btnConciliar.Enabled = false;
                    }

                    else
                    {
                        MessageBox.Show("Hubo un problema al Conciliar");
                    }

                }
                if (cboHojas.Text == "Maestro InBound Dolares")
                {
                    DataTable data = (DataTable)(dgvDatos.DataSource);

                    bool resultado = new Operaciones().cargarDatosEnWu(data);

                    if (resultado)
                    {
                        /////
                        //// NO CONCILIARON
                        ///
                        dgvDatos.Columns.Clear();
                        // SqlConnection cn = new SqlConnection(Conexion.cnn);
                        SqlCommand comandoRia = new SqlCommand("SELECT    t1.*,t2.jazc34idre as 'Id Remesa BT'," +
                            " t2.jazc34tip as 'Tipo (Env/Rec) BT', t2.jazc34fcc as 'Fecha BT',t2.jazc34mdp as 'Mda Pago BT'," +
                            " t2.jazc34imp as 'Importe BT',t2.jazc34ref as 'MTCN BT', t2.jazc34est as 'Estado BT'," +
                            " t2.JAZC34NDoO as 'NDoc Ordenante BT', t2.JAZC34NDoB as 'NDoc Beneficiario BT',t2.jazc34idER as 'EMPRESA BT' " +
                            "FROM conciliacionEnWu T1 FULL OUTER JOIN JAZC34 T2  " +
   "     ON  T1.MTCN = T2.jazc34ref   WHERE  " +
    "    ISNULL(CONVERT(varchar(12), CAST(T1.IMP_PAG AS decimal(20, 0))), -1) != ISNULL(CONVERT(varchar(12),  " +
   "     CAST(T2.[jazc34imp] AS decimal(20, 0))), -1) or  T1.MTCN != T2.jazc34ref  ", cn);
                        comandoRia.Parameters.AddWithValue("@fecha", selectDateAsString);
                        SqlDataAdapter adaptador = new SqlDataAdapter(comandoRia);
                        //  adaptador.SelectCommand = comandoRia;
                        DataTable tabla = new DataTable();
                        adaptador.Fill(tabla);
                        dgvDatos.DataSource = tabla;

                        //////
                        ///// CONCILIARION
                        ////
                        dgvIguales.Columns.Clear();
                        SqlCommand comandoIguales = new SqlCommand("SELECT CAST(t1.fecha AS varchar(10))  as 'Fecha', concat(T1.Apellido_Remitente, ' ', T1.Nombre_Remitente) as 'Remitente', " +
          "  t1.Beneficiario as 'Beneficiario', t1.MTCN as 'MTCN', t1.principal as 'Principal','0'as 'Cargo',t1.IMP_PAG ,'No aplica' as 'Total Enviado' ,t1.pais_Remitente, t1.ciudad_Remitente , t2.jazc34idER as 'EMPRESA' ,GETDATE()as 'Fecha Consulta'  " +
            "   FROM conciliacionEnWu T1 INNER JOIN JAZC34 T2   " +
   "     ON  T1.MTCN = T2.jazc34ref   WHERE   " +
    "    ISNULL(CONVERT(varchar(12), CAST(T1.IMP_PAG AS decimal(20, 0))), -1) = ISNULL(CONVERT(varchar(12), " +
   "     CAST(T2.[jazc34imp] AS decimal(20, 0))), -1)  " + " Drop table conciliacionEnWu", cn);
                        comandoIguales.Parameters.AddWithValue("@fecha", selectDateAsString);
                        SqlDataAdapter a = new SqlDataAdapter(comandoIguales);
                        //   a.SelectCommand = comandoIguales;
                        DataTable tabla2 = new DataTable();
                        a.Fill(tabla2);
                        dgvIguales.DataSource = tabla2;


                        MessageBox.Show("Se Concilio con Exito");
                        btnGuardar.Enabled = true;
                        btnConciliar.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("Hubo un problema al Conciliar");
                    }
                }
                if (cboHojas.Text == "BanTotal")
                {
                    DataTable data = (DataTable)(dgvDatos.DataSource);

                    bool resultado = new Operaciones().cargarBanTotal(data);

                   /* if (resultado)
                    {
                        /////
                        //// NO CONCILIARON
                        ///
                        dgvDatos.Columns.Clear();
                        // SqlConnection cn = new SqlConnection(Conexion.cnn);
                        SqlCommand comandoRia = new SqlCommand("select t1.* from Hoja1$ t1 inner join Remesa T2 on "+
                        " t1.[Id Remesa]=t2.IdTransferencia where  t2.Estado = 'RP'  and ISNULL(t2.CantidadEnvio, -1) != ISNULL(t1.importe, -1) and " +
                        " CAST(fecha AS varchar(10)) = CAST(fecha AS varchar(10))", cn);
                        SqlDataAdapter adaptador = new SqlDataAdapter();
                        adaptador.SelectCommand = comandoRia;
                        DataTable tabla = new DataTable();
                        adaptador.Fill(tabla);
                        dgvDatos.DataSource = tabla;

                        //////
                        ///// CONCILIARION
                        ////
                        dgvIguales.Columns.Clear();
                        SqlCommand comandoIguales = new SqlCommand("select t1.* from Hoja1$ t1 inner join Remesa T2 on " +
                        " t1.[Id Remesa]=t2.IdTransferencia where  t2.Estado = 'RP'  and ISNULL(t2.CantidadEnvio, -1) = ISNULL(t1.importe, -1) and " +
                        " CAST(fecha AS varchar(10)) = CAST(fecha AS varchar(10))"+ " drop table tableBantotal ", cn);
                        SqlDataAdapter a = new SqlDataAdapter();
                        a.SelectCommand = comandoIguales;
                        DataTable tabla2 = new DataTable();
                        a.Fill(tabla2);
                        dgvIguales.DataSource = tabla2;


                        MessageBox.Show("Se Concilio con Exito");
                        btnGuardar.Enabled = true;
                        btnConciliar.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("Hubo un problema al Conciliar");
                    }*/


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Verifiue los Datos" + ex.Message.ToString());
            }



        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnMostrar_Click(object sender, EventArgs e)
        {
            btnConciliar.Enabled = true;
            if (cboHojas.Text == "Sheet1")
            {
                btnMostrar.Enabled = false;
                
            }
           
           
            if (btnMostrar.Enabled)
            {
                btnConciliar.Enabled = true;
          

            }
            if (txtRuta.Text!= "CorrespPayment")
            {
            dgvDatos.DataSource = dt.Tables[cboHojas.SelectedIndex];
            btnMostrar.Text = "Borrar";
            for (int i = 0; i < 1; i++)
            {
                DataGridViewRow dgvDelRow = dgvDatos.Rows[i];
                dgvDatos.Rows.Remove(dgvDelRow);
                //  dgv/*Datos.Rows.RemoveAt(i);

            }
            }
            else
            {
                dgvDatos.DataSource = dt.Tables[cboHojas.SelectedIndex];
            }

        }

        private void btnGuardar_Click_1(object sender, EventArgs e)
        {
            dateTimePicker1.Enabled = true;
            DataTable d = (DataTable)(dgvDatos.DataSource);
            DataTable A = (DataTable)(dgvIguales.DataSource);
            bool r = new Operaciones().cargarBanTotal(d);
            bool b = new Operaciones().cargarBanTotal(A);
            SLDocument sl = new SLDocument();
          //  var r = ruta.Split('.');
            if (cboHojas.Text == "Sheet1")
            {

                SLStyle style = new SLStyle();
                style.Font.FontSize = 12;
                style.Font.Bold = true;

                int col = 1;
                foreach (DataGridViewColumn column in dgvDatos.Columns)
                {
                    sl.AddWorksheet("RIA- No Conciliarón");
                    sl.DeleteWorksheet("Sheet1");
                    sl.SelectWorksheet("Translations");
                    sl.SetCellValue(1, col, column.HeaderText.ToString());
                    sl.SetCellStyle(1, col, style);
                    col++;
                }

                int iR = 2;
                foreach (DataGridViewRow row in dgvDatos.Rows)
                {

                    sl.SetCellValue(iR, 1, row.Cells[0].Value.ToString());
                    sl.SetCellValue(iR, 2, row.Cells[1].Value.ToString());
                    sl.SetCellValue(iR, 3, row.Cells[2].Value.ToString());
                    sl.SetCellValue(iR, 4, row.Cells[3].Value.ToString());
                    sl.SetCellValue(iR, 5, row.Cells[4].Value.ToString());
                    sl.SetCellValue(iR, 6, row.Cells[5].Value.ToString());
                    sl.SetCellValue(iR, 7, row.Cells[6].Value.ToString());
                    sl.SetCellValue(iR, 8, row.Cells[7].Value.ToString());
                    sl.SetCellValue(iR, 9, row.Cells[8].Value.ToString());
                    sl.SetCellValue(iR, 10, row.Cells[9].Value.ToString());
                    sl.SetCellValue(iR, 11, row.Cells[10].Value.ToString());
                    sl.SetCellValue(iR, 12, row.Cells[11].Value.ToString());
                    sl.SetCellValue(iR, 13, row.Cells[12].Value.ToString());
                    sl.SetCellValue(iR, 14, row.Cells[13].Value.ToString());
                    sl.SetCellValue(iR, 15, row.Cells[14].Value.ToString());
                    sl.SetCellValue(iR, 16, row.Cells[15].Value.ToString());
                    sl.SetCellValue(iR, 17, row.Cells[16].Value.ToString());
                    sl.SetCellValue(iR, 18, row.Cells[17].Value.ToString());
                    sl.SetCellValue(iR, 19, row.Cells[18].Value.ToString());
                    sl.SetCellValue(iR, 20, row.Cells[19].Value.ToString());
                    sl.SetCellValue(iR, 21, row.Cells[20].Value.ToString());

                    iR++;
                }
                int col1 = 1;
                foreach (DataGridViewColumn column1 in dgvIguales.Columns)
                {

                    sl.AddWorksheet("RIA-Conciliarón");
                    sl.SetCellValue(1, col1, column1.HeaderText.ToString());
                    sl.SetCellStyle(1, col1, style);
                    col1++;
                }

                int iR1 = 2;
                foreach (DataGridViewRow row1 in dgvIguales.Rows)
                {

                    sl.SetCellValue(iR1, 1, row1.Cells[0].Value.ToString());
                    sl.SetCellValue(iR1, 2, row1.Cells[1].Value.ToString());
                    sl.SetCellValue(iR1, 3, row1.Cells[2].Value.ToString());
                    sl.SetCellValue(iR1, 4, row1.Cells[3].Value.ToString());
                    sl.SetCellValue(iR1, 5, row1.Cells[4].Value.ToString());
                    sl.SetCellValue(iR1, 6, row1.Cells[5].Value.ToString());
                    sl.SetCellValue(iR1, 7, row1.Cells[6].Value.ToString());
                    sl.SetCellValue(iR1, 8, row1.Cells[7].Value.ToString());
                    sl.SetCellValue(iR1, 9, row1.Cells[8].Value.ToString());
                    iR1++;

                }
            }
            if (cboHojas.Text == "Maestro OutBound Dolares")
            {


                SLStyle style = new SLStyle();
                style.Font.FontSize = 12;
                style.Font.Bold = true;

                int col = 1;
                foreach (DataGridViewColumn column in dgvDatos.Columns)
                {
                    sl.AddWorksheet("Western Envios-No Conciliarón");
                    sl.DeleteWorksheet("Sheet1");
                    sl.SelectWorksheet("Translations");
                    sl.SetCellValue(1, col, column.HeaderText.ToString());
                    sl.SetCellStyle(1, col, style);
                    col++;
                }

                int iR = 2;
                foreach (DataGridViewRow row in dgvDatos.Rows)
                {

                    sl.SetCellValue(iR, 1, row.Cells[0].Value.ToString());
                    sl.SetCellValue(iR, 2, row.Cells[1].Value.ToString());
                    sl.SetCellValue(iR, 3, row.Cells[2].Value.ToString());
                    sl.SetCellValue(iR, 4, row.Cells[3].Value.ToString());
                    sl.SetCellValue(iR, 5, row.Cells[4].Value.ToString());
                    sl.SetCellValue(iR, 6, row.Cells[5].Value.ToString());
                    sl.SetCellValue(iR, 7, row.Cells[6].Value.ToString());
                    sl.SetCellValue(iR, 8, row.Cells[7].Value.ToString());
                    sl.SetCellValue(iR, 9, row.Cells[8].Value.ToString());
                    sl.SetCellValue(iR, 10, row.Cells[9].Value.ToString());
                    sl.SetCellValue(iR, 11, row.Cells[10].Value.ToString());
                    sl.SetCellValue(iR, 12, row.Cells[11].Value.ToString());
                    sl.SetCellValue(iR, 13, row.Cells[12].Value.ToString());
                    sl.SetCellValue(iR, 14, row.Cells[13].Value.ToString());
                    sl.SetCellValue(iR, 15, row.Cells[14].Value.ToString());
                    sl.SetCellValue(iR, 16, row.Cells[15].Value.ToString());
                    sl.SetCellValue(iR, 17, row.Cells[16].Value.ToString());
                    sl.SetCellValue(iR, 18, row.Cells[17].Value.ToString());
                    sl.SetCellValue(iR, 19, row.Cells[18].Value.ToString());
                    sl.SetCellValue(iR, 20, row.Cells[19].Value.ToString());
                    sl.SetCellValue(iR, 21, row.Cells[20].Value.ToString());


                    iR++;
                }
                int col1 = 1;
                foreach (DataGridViewColumn column1 in dgvIguales.Columns)
                {

                    sl.AddWorksheet("Western Envios-Conciliarón");
                    sl.SetCellValue(1, col1, column1.HeaderText.ToString());
                    sl.SetCellStyle(1, col1, style);
                    col1++;
                }

                int iR1 = 2;
                foreach (DataGridViewRow row1 in dgvIguales.Rows)
                {

                    sl.SetCellValue(iR1, 1, row1.Cells[0].Value.ToString());
                    sl.SetCellValue(iR1, 2, row1.Cells[1].Value.ToString());
                    sl.SetCellValue(iR1, 3, row1.Cells[2].Value.ToString());
                    sl.SetCellValue(iR1, 4, row1.Cells[3].Value.ToString());
                    sl.SetCellValue(iR1, 5, row1.Cells[4].Value.ToString());
                    sl.SetCellValue(iR1, 6, row1.Cells[5].Value.ToString());
                    sl.SetCellValue(iR1, 7, row1.Cells[6].Value.ToString());
                    sl.SetCellValue(iR1, 8, row1.Cells[7].Value.ToString());
                    sl.SetCellValue(iR1, 9, row1.Cells[8].Value.ToString());
                    sl.SetCellValue(iR1, 10, row1.Cells[9].Value.ToString());
                    sl.SetCellValue(iR1, 11, row1.Cells[10].Value.ToString());
                    sl.SetCellValue(iR1, 12, row1.Cells[11].Value.ToString());
                 //   sl.SetCellValue(iR1, 13, row1.Cells[12].Value.ToString());

                    iR1++;
                }


            }
            if (cboHojas.Text == "Maestro InBound Dolares")
            {


                SLStyle style = new SLStyle();
                style.Font.FontSize = 12;
                style.Font.Bold = true;

                int col = 1;
                foreach (DataGridViewColumn column in dgvDatos.Columns)
                {
                    sl.AddWorksheet("Western Pagos-No Conciliarón");
                    sl.DeleteWorksheet("Sheet1");
                    sl.SelectWorksheet("Translations");
                    sl.SetCellValue(1, col, column.HeaderText.ToString());
                    sl.SetCellStyle(1, col, style);
                    col++;
                }

                int iR = 2;
                foreach (DataGridViewRow row in dgvDatos.Rows)
                {

                    sl.SetCellValue(iR, 1, row.Cells[0].Value.ToString());
                    sl.SetCellValue(iR, 2, row.Cells[1].Value.ToString());
                    sl.SetCellValue(iR, 3, row.Cells[2].Value.ToString());
                    sl.SetCellValue(iR, 4, row.Cells[3].Value.ToString());
                    sl.SetCellValue(iR, 5, row.Cells[4].Value.ToString());
                    sl.SetCellValue(iR, 6, row.Cells[5].Value.ToString());
                    sl.SetCellValue(iR, 7, row.Cells[6].Value.ToString());
                    sl.SetCellValue(iR, 8, row.Cells[7].Value.ToString());
                    sl.SetCellValue(iR, 9, row.Cells[8].Value.ToString());
                    sl.SetCellValue(iR, 10, row.Cells[9].Value.ToString());
                    sl.SetCellValue(iR, 11, row.Cells[10].Value.ToString());
                    sl.SetCellValue(iR, 12, row.Cells[11].Value.ToString());
                    sl.SetCellValue(iR, 13, row.Cells[12].Value.ToString());
                    sl.SetCellValue(iR, 14, row.Cells[13].Value.ToString());
                    sl.SetCellValue(iR, 15, row.Cells[14].Value.ToString());
                    sl.SetCellValue(iR, 16, row.Cells[15].Value.ToString());
                    sl.SetCellValue(iR, 17, row.Cells[16].Value.ToString());
                    sl.SetCellValue(iR, 18, row.Cells[17].Value.ToString());
                    sl.SetCellValue(iR, 19, row.Cells[18].Value.ToString());
                    sl.SetCellValue(iR, 20, row.Cells[19].Value.ToString());
                    sl.SetCellValue(iR, 21, row.Cells[20].Value.ToString());

                    iR++;
                }
                int col1 = 1;
                foreach (DataGridViewColumn column1 in dgvIguales.Columns)
                {

                    sl.AddWorksheet("Western Pagos-Conciliarón");
                    sl.SetCellValue(1, col1, column1.HeaderText.ToString());
                    sl.SetCellStyle(1, col1, style);
                    col1++;
                }

                int iR1 = 2;
                foreach (DataGridViewRow row1 in dgvIguales.Rows)
                {

                    sl.SetCellValue(iR1, 1, row1.Cells[0].Value.ToString());
                    sl.SetCellValue(iR1, 2, row1.Cells[1].Value.ToString());
                    sl.SetCellValue(iR1, 3, row1.Cells[2].Value.ToString());
                    sl.SetCellValue(iR1, 4, row1.Cells[3].Value.ToString());
                    sl.SetCellValue(iR1, 5, row1.Cells[4].Value.ToString());
                    sl.SetCellValue(iR1, 6, row1.Cells[5].Value.ToString());
                    sl.SetCellValue(iR1, 7, row1.Cells[6].Value.ToString());
                    sl.SetCellValue(iR1, 8, row1.Cells[7].Value.ToString());
                    sl.SetCellValue(iR1, 9, row1.Cells[8].Value.ToString());
                    sl.SetCellValue(iR1, 10, row1.Cells[9].Value.ToString());
                    sl.SetCellValue(iR1, 11, row1.Cells[10].Value.ToString());

                    iR1++;
                }


            }
          

            /*  int col2 = 1;
              foreach (DataGridViewColumn column2 in dgvFaltan.Columns)
              {
                  sl.AddWorksheet("No se Encuentran");
                  sl.SetCellValue(1, col2, column2.HeaderText.ToString());
                  sl.SetCellStyle(1, col2, style);
                  col2++;
              }

                int iR2 = 2;
              foreach (DataGridViewRow row2 in dgvFaltan.Rows)
              {

                  sl.SetCellValue(iR2, 1, row2.Cells[0].Value.ToString());
                  sl.SetCellValue(iR2, 2, row2.Cells[1].Value.ToString());
                  sl.SetCellValue(iR2, 3, row2.Cells[2].Value.ToString());
                  sl.SetCellValue(iR2, 4, row2.Cells[3].Value.ToString());
                  sl.SetCellValue(iR2, 5, row2.Cells[4].Value.ToString());
                  sl.SetCellValue(iR2, 6, row2.Cells[5].Value.ToString());
                  sl.SetCellValue(iR2, 7, row2.Cells[6].Value.ToString());
                  sl.SetCellValue(iR2, 8, row2.Cells[7].Value.ToString());
                  sl.SetCellValue(iR2, 9, row2.Cells[8].Value.ToString());
                  sl.SetCellValue(iR2, 10, row2.Cells[9].Value.ToString());
                  sl.SetCellValue(iR2, 11, row2.Cells[10].Value.ToString());

                  iR++;
               }*/



            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Guardar archivo";
            saveFileDialog1.CheckPathExists = true;
            saveFileDialog1.DefaultExt = "xlsx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    sl.SaveAs(saveFileDialog1.FileName);
                    btnImportar.Enabled = true;                  
                  //  chkbantotal.Checked = false;
                   // btnBantotal.Enabled = false;
                    dateTimePicker1.Enabled = true;
                    
                    MessageBox.Show("Archivo exportado con éxito");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            // sl.SaveAs(@"C:\Users\jose_\Desktop\destinoExcel\conciliacion.xlsx");
        }

        private void dgvFaltan_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void cboHojas_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            if (cboHojas.Text == "Sheet1")
            {
                btnMostrar.Enabled = true;
                btnMostrar.Text = "Mostrar";
                lbltitulo.Text = "DATOS RIA";
            }
            if (cboHojas.Text == "Maestro OutBound Dolares")
            {
                btnMostrar.Text = "Mostrar";
                btnMostrar.Enabled = true;
                lbltitulo.Text = "DATOS Western Envios";
            }
            if (cboHojas.Text == "Maestro InBound Dolares")
            {
                btnMostrar.Text = "Mostrar";
                btnMostrar.Enabled = true;
                lbltitulo.Text = "DATOS Western Pagos";
            }
        }

        private void dgvIguales_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgvDatos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void chkUrl_CheckedChanged(object sender, EventArgs e)
        {
           // btnBantotal.Enabled = true;
            dateTimePicker1.Enabled = true;
          
                btnImportar.Enabled = false;
                cboHojas.Items.Clear();
                btnGuardar.Enabled = false;
                txtRuta.Text = "";
                btnMostrar.Enabled = false;
                btnConciliar.Enabled = false;
                dgvDatos.Columns.Clear();
            /////
                btnImportar.Enabled = true;
              //  btnBantotal.Enabled = false;
                dateTimePicker1.Enabled = false;
                cboHojas.Items.Clear();
                btnGuardar.Enabled = false;
                txtRuta.Text = "";
                btnMostrar.Enabled = false;
                btnConciliar.Enabled = false;
                dgvDatos.Columns.Clear();
            
        }

        private void txtUrl_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnBantotal_Click(object sender, EventArgs e)
        {
            string selectDateAsString = dateTimePicker1.Value.ToString("yyyy-MM-dd");
           
          
            SqlConnection cn = new SqlConnection(Conexion.cnn);

                         
                //////
                ///// Data
                ////
                dgvDatos.Columns.Clear();
            string comandoIguales = "SELECT [Id Remesa],[Tipo (Env/Rec)],CONVERT(nvarchar(10),[Fecha]) 'Fecha',[Mda Pago],[Importe],RTRIM([MTCN]) 'MTCN',[Estado],RTRIM([NDoc Ordenante]) 'NDoc Ordenante', " +
                    " RTRIM([NDoc Beneficiario]) 'NDoc Beneficiario'  FROM[Hoja1$] where CAST(fecha AS varchar(10)) = @fecha"; //selectDateAsString.ToString(), cn);
            /*new SqlCommand("select Id_Remesa as 'Id Remesa', Tipo_Env_Rec as 'Tipo (Env/Rec)', "+
" Fecha as 'Fecha', Mda_Pago as 'Mda Pago', Importe as 'Importe', "+
" MTCN as MTCN, Estado as 'Estado', NDoc_Ordenante as 'NDoc Ordenante', "+
" NDoc_Beneficiario as 'NDoc Beneficiario'   from DataCon", cn);// + " drop table conciliacionEnWu ", cn);*/
            SqlCommand cmd = new SqlCommand(comandoIguales, cn);
            cmd.Parameters.AddWithValue("@fecha", selectDateAsString );

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable data = new DataTable();
            da.Fill(data);

            da.Dispose();
            cn.Close();

            dgvDatos.DataSource = data;
            cboHojas.Items.Insert(0,"BanTotal");           
            btnConciliar.Enabled = true;
            MessageBox.Show("Consulta Realizado");
             
           
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
        }
    }
}