

namespace EKPolizasSemp
{


    #region Libraries (Librerias)
    using System;
    using System.Data;
    using System.Data.SqlClient;
    using System.IO;
    using System.Windows.Forms;
    using EDsemp.Classes;
    using DevComponents.DotNetBar;  
    #endregion
    public partial class Form1 : Office2007Form
    {
        #region Properties (proiedades)
        string LOCALIDAD;
        string CUENTA_CAJA;
        string CUENTA_INTERES;
        string CUENTA_IVA;
        string NUMERO_DE_CAJA;
        string DEPARTAMENTO;
        string CUENTA_PRESTAMOS;
        string TABLA_CAJA;
        string LUGAR_CONTA;
        string empresa_Conta;
        string mes_letra;
        string a, b, c, f;//variables de conexion al servidor
        string sqlcnx;
        string base_de_datos;
        string logotipo;
        string path;
        //
        string año = "";
        string mes = "";
        int nveces;
        int dias_en_mes;
        string fecha_inicial;
        string fecha_final;
        string fecha_dos;
        string fecha_uno;
        string fecha_tres;
        string fecha_cuatro;
        string fecha_cinco;
        #endregion


        #region Methods (Metodos)

        //llenos los combos
        private void combos()
        {
            this.comboBoxEx1.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.comboBoxEx1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.comboBoxEx1.DropDownStyle = ComboBoxStyle.DropDownList;

            this.comboBoxEx2.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.comboBoxEx2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.comboBoxEx2.DropDownStyle = ComboBoxStyle.DropDownList;

            this.comboBoxEx3.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.comboBoxEx3.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.comboBoxEx3.DropDownStyle = ComboBoxStyle.DropDownList;

            this.comboBoxEx4.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.comboBoxEx4.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.comboBoxEx4.DropDownStyle = ComboBoxStyle.DropDownList;


            string[] meses;
            meses = new string[] { "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE" };//PRA DECLARAR UNA MATRIZ TIPO CADENA
            this.comboBoxEx2.DataSource = meses;

            string[] año;
            año = new String[] { "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024" };
            this.comboBoxEx3.DataSource = año;


            //combo de loacalidades
            SqlConnection conexionmy = new SqlConnection();
            conexionmy.ConnectionString = sqlcnx;//conexion mysql
            conexionmy.Open();
            DataTable tablasql = new DataTable();
            SqlDataAdapter datosSql = new SqlDataAdapter("USE SEMP2013_NAU_1 " +
                "Select [Nombre Sucursal] from Localidades where  CONCEPTO='CASA DE EMPEÑO' order by No asc " +
                "", conexionmy);
            datosSql.Fill(tablasql);
            comboBoxEx1.ValueMember = "Nombre Sucursal";
            comboBoxEx1.DisplayMember = "Nombre Sucursal";
            comboBoxEx1.DataSource = tablasql;
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            conexion();
        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            circularProgress1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            circularProgress1.Visible = false;
            circularProgress1.IsRunning = false;
            combos();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {

        }

       

        private void buttonX1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialogoRuta = new FolderBrowserDialog();

            if (dialogoRuta.ShowDialog() == DialogResult.OK)
            {

                path= dialogoRuta.SelectedPath;
                textBoxX1.Text = path;

            }
        }

        //conexion de generador de polizas
        private void conexion()
        {
            //aligual que las demas aplicaciones cargaremos nuestra llave al servidor de oficinas para la conexion directa
            string cadena = "C:/SEMP2013/GeneradorPolizasSEMP/Polizas_General/Polizas_General/bin/Debug/cdblista.txt";
            using (StreamReader sr1 = new StreamReader(cadena, true))
            {

                string lineA = sr1.ReadLine();
                string lineB = sr1.ReadLine();
                string lineC = sr1.ReadLine();
                string lineF = sr1.ReadLine();
                
                //ahroa desecrypto la informacion             
                a = Encriptar_Desencriptar.DecryptKeyMD5(lineA);
                b = Encriptar_Desencriptar.DecryptKeyMD5(lineB);
                c = Encriptar_Desencriptar.DecryptKeyMD5(lineC);  
                f = Encriptar_Desencriptar.DecryptKeyMD5(lineF);
                //ahora realizo la conexion par amostrar las sucursales
             
                try
                {

                    sqlcnx = "Data Source=" + a + " ;" +
                        "Initial Catalog=" + b + ";" +
                        "Persist Security Info=True;" +
                        "User ID=" + c + ";Password=" + f + "";
                    SqlConnection conexion = new SqlConnection();
                    conexion.ConnectionString = sqlcnx;
                    conexion.Open();

                    if (true)
                    {
                        ////envio al form de reportes
                        //combos();
                    }
                    else
                    {
                        MessageBox.Show("Error de Conexion","Polizas Semp",MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }






            }

        }

        #endregion

        #region MyRegion

        #endregion

        #region MyRegion

        #endregion

        #region MyRegion

        #endregion

        public Form1()
        {
            InitializeComponent();
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            circularProgress1.Visible = true;
            circularProgress1.IsRunning = true;
            backgroundWorker1.RunWorkerAsync();
        }
    }
}
