using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Conciliacion.DATA
{
   public class Operaciones
    {
        public bool cargarDatos(DataTable tbData)
        {
            bool resultado = true;
            using (SqlConnection cn= new SqlConnection(Conexion.cnn))
            {
                cn.Open();
                SqlCommand comandoRia = new SqlCommand("TablaConciliacionRia", cn);
                comandoRia.CommandType = CommandType.StoredProcedure;
                comandoRia.ExecuteNonQuery();
              
                using (SqlBulkCopy s = new SqlBulkCopy(cn))
                {
                    
                    //  asignamos (columnas de origen, columnas destino)
                    s.ColumnMappings.Add("Date", "fecha");
                    s.ColumnMappings.Add("Location", "ubicacion");
                    s.ColumnMappings.Add("Item", "articulo");
                    s.ColumnMappings.Add("Beneficiary", "beneficiario");
                    s.ColumnMappings.Add("PIN", "pin");
                    s.ColumnMappings.Add("Seq", "Seq");
                    s.ColumnMappings.Add("Cur", "tipo_moneda");
                    s.ColumnMappings.Add("Order Amt", "importe_pago");
                    s.ColumnMappings.Add("Order Amt USD", "importe_Pago_USD");
                    s.ColumnMappings.Add("Commission Amt", "Comision");
                    s.ColumnMappings.Add("Commission Amt USD", "comision_dolares");

                    //tabla de destino
                   
                    s.DestinationTableName = "conciliacion";
                    try
                    {
                        s.WriteToServer(tbData);
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        resultado = false;
                    }
                }


            }

                return resultado;
        }

        public bool cargarDatosWu(DataTable tbData)
        {
            bool resultado = true;
            using (SqlConnection cn = new SqlConnection(Conexion.cnn))
            {
                cn.Open();
                SqlCommand comandoRia = new SqlCommand("TablaConciliacionWu", cn);
                comandoRia.CommandType = CommandType.StoredProcedure;
                comandoRia.ExecuteNonQuery();

                using (SqlBulkCopy s = new SqlBulkCopy(cn))
                {

                    //  asignamos (columnas de origen, columnas destino)
                    s.ColumnMappings.Add("Column0", "Agencia");
                    s.ColumnMappings.Add("Column1", "Fecha");
                    s.ColumnMappings.Add("Column2", "Apellido_Beneficiario");
                    s.ColumnMappings.Add("Column3", "Nombre_Beneficiario");
                    s.ColumnMappings.Add("Column4", "Remitente");
                    s.ColumnMappings.Add("Column5", "MTCN");
                    s.ColumnMappings.Add("Column6", "Principal");
                    s.ColumnMappings.Add("Column7", "Cargo");
                    s.ColumnMappings.Add("Column8", "IGV");
                    s.ColumnMappings.Add("Column9", "ITF");
                    s.ColumnMappings.Add("Column10", "Total_Enviado");
                    s.ColumnMappings.Add("Column11", "Pais_Beneficiario");
                    s.ColumnMappings.Add("Column12", "Ciudad_Beneficiario");

                    /* s.ColumnMappings.Add("Agencia", "Agencia");
                     s.ColumnMappings.Add("Fecha", "Fecha");
                     s.ColumnMappings.Add("Apellido Beneficiario", "Apellido_Beneficiario");
                     s.ColumnMappings.Add("Nombre Beneficiario", "Nombre_Beneficiario");
                     s.ColumnMappings.Add("Remitente", "Remitente");
                     s.ColumnMappings.Add("MTCN", "MTCN");
                     s.ColumnMappings.Add("Principal", "Principal");
                     s.ColumnMappings.Add("Cargo", "Cargo");
                     s.ColumnMappings.Add("I.G.V.", "IGV");
                     s.ColumnMappings.Add("I.T.F.", "ITF");
                     s.ColumnMappings.Add("Total Enviado", "Total_Enviado");
                     s.ColumnMappings.Add("País Beneficiario", "Pais_Beneficiario");
                     s.ColumnMappings.Add("Ciudad Beneficiario", "Ciudad_Beneficiario");*/



                    //tabla de destino
                    s.DestinationTableName = "conciliacionWu";
                    try
                    {
                        s.WriteToServer(tbData);
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        resultado = false;
                    }
                }


            }

            return resultado;
        }

        public bool cargarDatosEnWu(DataTable tbData)
        {
            bool resultado = true;
            using (SqlConnection cn = new SqlConnection(Conexion.cnn))
            {
                cn.Open();
                SqlCommand comandoRia = new SqlCommand("TablaConciliacionEnWu", cn);
                comandoRia.CommandType = CommandType.StoredProcedure;
                comandoRia.ExecuteNonQuery();

                using (SqlBulkCopy s = new SqlBulkCopy(cn))
                {

                    //  asignamos (columnas de origen, columnas destino)
                    s.ColumnMappings.Add("Column0", "Agencia");
                    s.ColumnMappings.Add("Column1", "Fecha");
                    s.ColumnMappings.Add("Column2", "Apellido_Remitente");
                    s.ColumnMappings.Add("Column3", "Nombre_Remitente");
                    s.ColumnMappings.Add("Column4", "Beneficiario");
                    s.ColumnMappings.Add("Column5", "MTCN"); ;
                    s.ColumnMappings.Add("Column6", "IMP_PAG");
                    s.ColumnMappings.Add("Column7", "ITF");
                    s.ColumnMappings.Add("Column8", "Principal");
                    s.ColumnMappings.Add("Column9", "Pais_Remitente");
                    s.ColumnMappings.Add("Column10", "Ciudad_Remitente");

                  


                    //tabla de destino
                    s.DestinationTableName = "conciliacionEnWu";
                    try
                    {
                        s.WriteToServer(tbData);
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        resultado = false;
                    }
                }


            }

            return resultado;
        }

        public bool cargarBanTotal(DataTable tbData)
        {

            bool resultado = true;
            using (SqlConnection cn = new SqlConnection(Conexion.cnn))
            {
                cn.Open();
              /*      SqlCommand comandoRia = new SqlCommand("TablaConciliacionBantotal", cn);
                    comandoRia.CommandType = CommandType.StoredProcedure;
                    //comandoRia.ExecuteNonQuery();*/

                using (SqlBulkCopy s = new SqlBulkCopy(cn))
                {

                    //  asignamos (columnas de origen, columnas destino)
                    s.ColumnMappings.Add("Fecha", "Fecha");
                    s.ColumnMappings.Add("Beneficiario", "Beneficiario");
                    s.ColumnMappings.Add("Remitente", "Remitente");
                    s.ColumnMappings.Add("MTCN", "mtcn");
                    s.ColumnMappings.Add("Principal", "principal");
                    s.ColumnMappings.Add("Cargo", "Cargo");
                    s.ColumnMappings.Add("Total Enviado", "Total Pago");
                    s.ColumnMappings.Add("EMPRESA", "EMPRESA");
                    s.ColumnMappings.Add("Fecha Consulta", "Fecha Consulta");


                    //tabla de destino
                    s.DestinationTableName = "reporteConciliacion";
                    try
                    {
                        s.WriteToServer(tbData);
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        resultado = false;
                    }
                }


            }

            return resultado;
        }
        public bool cargarTabla(DataTable tbData)
        {

            bool resultado = true;
            using (SqlConnection cn = new SqlConnection(Conexion.cnn))
            {
                cn.Open();
                    SqlCommand comandoRia = new SqlCommand("tablaBantotal", cn);
                    comandoRia.CommandType = CommandType.StoredProcedure;
                   comandoRia.ExecuteNonQuery();

                using (SqlBulkCopy s = new SqlBulkCopy(cn))
                {

                    //  asignamos (columnas de origen, columnas destino)
                    s.ColumnMappings.Add("Id Remesa", "jazc34idre");
                    s.ColumnMappings.Add("Tipo (Env/Rec)", "jazc34tip");                   
                    s.ColumnMappings.Add("Mda Pago", "jazc34mdp");                    
                    s.ColumnMappings.Add("Importe", "jazc34imp");
                    s.ColumnMappings.Add("MTCN", "jazc34ref");
                    s.ColumnMappings.Add("Empresa", "jazc34idER");
                    s.ColumnMappings.Add("Estado", "jazc34est");
                    s.ColumnMappings.Add("Fecha", "jazc34fcc");
                    s.ColumnMappings.Add("NDoc ordenante", "JAZC34NDoO");
                    s.ColumnMappings.Add("NDoc Beneficiario", "JAZC34NDoB");


                    //tabla de destino
                    s.DestinationTableName = "JAZC34";
                    try
                    {
                        s.WriteToServer(tbData);
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        resultado = false;
                    }
                }


            }

            return resultado;
        }

    }
}
