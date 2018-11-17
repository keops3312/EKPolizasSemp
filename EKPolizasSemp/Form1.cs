﻿

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
    using System.Collections.Generic;
    using System.Linq;
    using DevComponents.DotNetBar.Controls;
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
        int nveces=0;
        int dias_en_mes;
        string fecha_inicial;
        string fecha_final;
        string fecha_dos;
        string fecha_uno;
        string fecha_tres;
        string fecha_cuatro;
        string fecha_cinco;

        string caja;
        string server;
        string letra;
        int opcion;


        #endregion

        #region Methods (Metodos)
        public Form1()
        {
            InitializeComponent();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            backgroundWorker2.WorkerReportsProgress = true;
            backgroundWorker2.WorkerSupportsCancellation = true;
            backgroundWorker3.WorkerReportsProgress = true;
            backgroundWorker3.WorkerSupportsCancellation = true;
            backgroundWorker4.WorkerReportsProgress = true;
            backgroundWorker4.WorkerSupportsCancellation = true;
        }

        //find place
        private void paso2()
        {

            //1. busco la tabla de datos en el servidor segun la sucursal seleccionada

            try
            {
                SqlConnection conexionmy = new SqlConnection(sqlcnx);
                DataTable tablasql = new DataTable();
                SqlDataAdapter datosSql = new SqlDataAdapter("USE SEMP2013_NAU_1 " +
                    "Select Logotipo, BD, lugar_conta, Empresa  from Localidades where  [Nombre Sucursal]='" + this.comboBoxEx1.Text + "'" +
                    "", conexionmy);
                datosSql.Fill(tablasql);
                base_de_datos = tablasql.Rows[0].ItemArray[1].ToString();
                LUGAR_CONTA = tablasql.Rows[0].ItemArray[2].ToString();
                logotipo = tablasql.Rows[0].ItemArray[0].ToString();
                empresa_Conta = tablasql.Rows[0].ItemArray[3].ToString();
                pictureBox1.Load(logotipo);
                this.label4.Text = base_de_datos;



                //2. cargo las cajas de la sucursal seleccionada

                DataTable tablecaja_dos = new DataTable();
                SqlDataAdapter datoscaja = new SqlDataAdapter("USE " + this.label4.Text + " " +
                    " Select NumCaja from selcaja order by NumCaja asc", conexionmy);
                datoscaja.Fill(tablecaja_dos);
                comboBoxEx4.ValueMember = "NumCaja";
                comboBoxEx4.DisplayMember = "NumCaja";
                comboBoxEx4.DataSource = tablecaja_dos;


                this.dataGridView5.DataSource = tablecaja_dos;



            }
            catch (Exception ex)
            {
                MessageBox.Show("No encontre datos de localidad intenta con una diferente por favor");
            }


         
        }

        //Fill Comboboxes
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
        //Range Dates
        private void mes_calculo()
        {
            //1obtener la fecha inicial

            switch (this.comboBoxEx2.Text)
            {
                case "ENERO":
                    mes = "01";
                    break;
                case "FEBRERO":
                    mes = "02";
                    break;
                case "MARZO":
                    mes = "03";
                    break;
                case "ABRIL":
                    mes = "04";
                    break;
                case "MAYO":
                    mes = "05";
                    break;
                case "JUNIO":
                    mes = "06";
                    break;
                case "JULIO":
                    mes = "07";
                    break;
                case "AGOSTO":
                    mes = "08";
                    break;
                case "SEPTIEMBRE":
                    mes = "09";
                    break;
                case "OCTUBRE":
                    mes = "10";
                    break;
                case "NOVIEMBRE":
                    mes = "11";
                    break;
                case "DICIEMBRE":
                    mes = "12";
                    break;
                default:
                    break;

            }
            año = this.comboBoxEx3.Text;
            dias_en_mes = DateTime.DaysInMonth(Convert.ToInt32(año), Convert.ToInt32(mes));
            //2 obtener la fecha final de cada mes
            fecha_inicial = año + "-" + mes + "-" + "01 00:00:00";
            fecha_final = año + "-" + mes + "-" + dias_en_mes + " 00:00:00";


            //obtengo datos de la caja
            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            DataTable tablasql = new DataTable();
            SqlDataAdapter datosSql = new SqlDataAdapter("USE " + this.label4.Text + "  " +
                "Select *  from contabilidad where leyenda2='" + this.comboBoxEx4.Text + "'" +
                "", conexionmy);
            datosSql.Fill(tablasql);
            LOCALIDAD = tablasql.Rows[0].ItemArray[2].ToString();
            CUENTA_CAJA = tablasql.Rows[0].ItemArray[6].ToString();
            CUENTA_INTERES = tablasql.Rows[0].ItemArray[7].ToString();
            CUENTA_IVA = tablasql.Rows[0].ItemArray[8].ToString();
            NUMERO_DE_CAJA = tablasql.Rows[0].ItemArray[3].ToString();
            DEPARTAMENTO = tablasql.Rows[0].ItemArray[9].ToString();
            CUENTA_PRESTAMOS = tablasql.Rows[0].ItemArray[5].ToString();
            TABLA_CAJA = tablasql.Rows[0].ItemArray[11].ToString();

        }

        #endregion

        #region Events (eventos)
        private void Form1_Load(object sender, EventArgs e)
        {
            circularProgress1.Visible = true;
            circularProgress1.IsRunning = true;
            backgroundWorker1.RunWorkerAsync();
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
            DialogResult pregunta = MessageBox.Show("Generar Polizas?", "EKPolizasSemp", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (pregunta == DialogResult.Yes)
            {
                //obtener los rangos, fechas de cada mes
                caja = this.comboBoxEx4.Text;
                server = this.label4.Text;
                path = this.textBoxX1.Text;

                mes_calculo();
             
                crear_poliza();
                
            }
        }

        private void crear_poliza()
        {
            //Todos los checkbox serán considerados como un item de checkboxlist:
            List<CheckBoxX> chkLst = new List<CheckBoxX>();

            //Añadimos todos los checkboxes al checkboxlist:
            chkLst.AddRange(this.groupPanel1.Controls.OfType<CheckBoxX>());

            for (int i = 0; i < chkLst.Count; i++)
            {

                if(chkLst[i].Checked==true && chkLst[i].Text == "Prestamos")
                {
                    opcion = 1;
                    letra = this.textBoxX3.Text;
                    backgroundWorker2.RunWorkerAsync();
                }

                if (chkLst[i].Checked == true && chkLst[i].Text == "Cobros")
                {
                    opcion = 2;
                    letra = this.textBoxX2.Text;
                    backgroundWorker3.RunWorkerAsync();
                }

                if (chkLst[i].Checked == true && chkLst[i].Text == "Diario")
                {
                    opcion = 3;
                    letra = this.textBoxX4.Text;
                    backgroundWorker4.RunWorkerAsync();
                }


                if (chkLst[i].Checked == true && chkLst[i].Text == "Interes Semanal")
                {
                   
                }

                if (chkLst[i].Checked == true && chkLst[i].Text == "Remision Semanal")
                {
                 
                }


                //Hacemos el Deschecked a todos los checkbox:
                // chkLst[i].Checked = true;

                //Si quisieramos que todos los chekbox estuvieran en checked:
                //chkLst[i].Checked = true

                //Si quisieramos que todos los chekbox se habilitaran:
                //chkLst[i].Enabled = true

                //Si quisieramos que todos los chekbox se deshabilitaran:
                //chkLst[i].Enabled = false


            }

        }


        #region Prestamos rutina
        private void prestamos()
        {
            try
            {

                progressBarX1.Maximum = dias_en_mes;
                progressBarX1.Value = 0;
                //string CONTRATO, STATUS, PRESTAMO;
                for (nveces = 1; nveces <= dias_en_mes; nveces++)
                {

                    backgroundWorker2.ReportProgress(nveces);

                    //2. cargo las cajas de la sucursal seleccionada
                    SqlConnection conexionmy = new SqlConnection();
                    conexionmy.ConnectionString = sqlcnx;//conexion mysql
                    conexionmy.Open();
                    SqlCommand comando_crea = new SqlCommand("USE " + server +
                    " truncate table prestamos_poliza " +
                    " ", conexionmy);
                    comando_crea.ExecuteNonQuery();
                    conexionmy.Close();

                    fecha_cinco = Convert.ToString(nveces);
                    if (fecha_cinco.Length == 1)
                    {
                        fecha_cinco = "0" + fecha_cinco;
                    }
                    fecha_uno = año + "-" + mes + "-" + fecha_cinco;
                    //MessageBox.Show("" + fecha_cinco);
                    fecha_dos = fecha_cinco + "/" + mes + "/" + año;

                    conexionmy.Open();
                    SqlCommand simbolo = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + letra + "')", conexionmy);
                    simbolo.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand diadepoliza = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + fecha_cinco + "')", conexionmy);
                    diadepoliza.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "PRESTAMOS DEL " + fecha_dos + " " + NUMERO_DE_CAJA + "')", conexionmy);
                    leyendaP.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + LOCALIDAD + " " + "CAJA # " + NUMERO_DE_CAJA + "')", conexionmy);
                    leyendaPL.ExecuteNonQuery();
                    conexionmy.Close();

                    poliza_prestamos();
                }//for 


            }//try
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void poliza_prestamos()
        {


            SqlConnection conexionmy = new SqlConnection();
            conexionmy.ConnectionString = sqlcnx;//conexion mysql


            fecha_tres = año;
            fecha_cuatro = mes;

            fecha_dos = fecha_cinco + "/" + mes + "/" + año;
            int año_inicial, nom, valor;
            string seis = "";
            string cuatro;
            string catorce;
            año_inicial = 2006;
            nom = 1;
            valor = Convert.ToInt32(año);
            cuatro = "0" + mes;
            while (año_inicial < valor)
            {
                nom = nom + 1;
                año_inicial = año_inicial + 1;

            }
            if (nom < 10)
            {
                seis = Convert.ToString("00" + nom);
            }
            if (nom >= 10)
            {
                seis = Convert.ToString("0" + nom);
            }

            catorce = CUENTA_PRESTAMOS + "-" + seis + "-" + cuatro; //cuanta para prestamos
            //la consulta
            DataTable tablapres = new DataTable();
            tablapres.Clear();
            SqlDataAdapter datos_pres = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  contratos INNER JOIN " + caja + " " +
                " ON contratos.contrato = " + caja + ".contrato  " +
                " WHERE contratos.Fechacons= '" + Convert.ToDateTime(fecha_uno).ToString("dd-MM-yyyy") + "' and " + caja + ".concepto LIKE '%PRESTAMO%' " +
                " order by contratos.contrato asc " +
                "", conexionmy);
            datos_pres.Fill(tablapres);
            //this.dataGridView1.DataSource = tablapres;
            //veo si esta vacio el dia
            if (tablapres.Rows.Count == 0)
            {
                //exportar con nota de no hubo operacion este dia y seguir con rutina
            }
            else
            {
                string CONTRATO, STATUS, PRESTAMO;
                foreach (DataRow dr in tablapres.Rows)
                {

                    CONTRATO = dr[1].ToString();//contrato
                    STATUS = dr[8].ToString(); //status
                    PRESTAMO = dr[13].ToString(); //prestamo $$

                    //NUMERO DE LA CUENTA
                    conexionmy.Open();
                    SqlCommand comando_A = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + catorce + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    comando_A.ExecuteNonQuery();
                    conexionmy.Close();
                    //LAS LEYENDAS
                    conexionmy.Open();
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "PRESTAMO CONTRATO # " + CONTRATO + "')", conexionmy);
                    comando_B.ExecuteNonQuery();
                    conexionmy.Close();
                    //EL MONTO DEL PRESTAMO
                    string monto_valor;

                    decimal unosidebedepagare = decimal.Round(Convert.ToDecimal(PRESTAMO), 2, MidpointRounding.AwayFromZero);
                    monto_valor = string.Format("{0:0.000000}", unosidebedepagare);


                    if (STATUS == "CANCELADO")
                    {
                        conexionmy.Open();
                        SqlCommand comando_C = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        comando_C.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_D = new SqlCommand("USE " + this.label4.Text + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + monto_valor + ",1.00" + "')", conexionmy);
                        comando_D.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    //AHORA LA CUENTA DE LA CAJA
                    conexionmy.Open();
                    SqlCommand comando_E = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + CUENTA_CAJA + "                ," + "  " + DEPARTAMENTO + " ')", conexionmy);
                    comando_E.ExecuteNonQuery();
                    conexionmy.Close();
                    //ahora leyenda caja
                    conexionmy.Open();
                    SqlCommand comando_F = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "CAJA" + " " + NUMERO_DE_CAJA + "CONTRATO #  " + CONTRATO + "')", conexionmy);
                    comando_F.ExecuteNonQuery();
                    conexionmy.Close();
                    //inserto espacio
                    //ahora la el monto de prestamo
                    conexionmy.Open();
                    SqlCommand comando_G = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('')", conexionmy);
                    comando_G.ExecuteNonQuery();
                    conexionmy.Close();
                    //nuevamene la cantidad
                    //ahora la el monto de prestamo             
                    string h1;
                    string valor_convertido;

                    decimal unosidebedepagares = decimal.Round(Convert.ToDecimal(PRESTAMO), 2, MidpointRounding.AwayFromZero);
                    valor_convertido = string.Format("{0:0.000000}", unosidebedepagares);
                    if (STATUS == "CANCELADO")
                    {
                        conexionmy.Open();
                        SqlCommand comando_H = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        comando_H.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_I = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + valor_convertido + ",1.00" + "')", conexionmy);
                        comando_I.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                }//cierre del ciclo for each
                conexionmy.Open();
                SqlCommand comando_J = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('FIN')", conexionmy);
                comando_J.ExecuteNonQuery();
                conexionmy.Close();
                notas();

            }//if del rows


            //creamos la poliza en la ruta del primer dia

        }


        private void notas()
        {
            SqlConnection conexionmy = new SqlConnection();
            conexionmy.ConnectionString = sqlcnx;//conexion mysql
            DataTable tablaprestamos = new DataTable();
            tablaprestamos.Clear();
            SqlDataAdapter datos_presa = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  prestamos_poliza " +
                " order by  no asc " +
                "", conexionmy);
            datos_presa.Fill(tablaprestamos);
            //this.dataGridView2.DataSource = tablaprestamos;
            System.IO.StreamWriter escribir = new System.IO.StreamWriter(path + "/PRESTAMOS " + caja + "-" + fecha_uno + ".pol");
            foreach (DataRow fila in tablaprestamos.Rows)
                escribir.WriteLine(fila[1]);
            escribir.Close();

        }

        #endregion



        #region Cobros rutina
        public void cobros()
        {
            try
            {

                //2. cargo las cajas de la sucursal seleccionada
                SqlConnection conexionmy = new SqlConnection(sqlcnx);
               
                progressBarX2.Maximum = dias_en_mes;
                progressBarX2.Value = 0;
                //string CONTRATO, STATUS, PRESTAMO;
                for (int nveces = 1; nveces <= dias_en_mes; nveces++)
                    {
                    backgroundWorker3.ReportProgress(nveces);

                    fecha_cinco = Convert.ToString(nveces);
                        if (fecha_cinco.Length == 1)
                        {
                            fecha_cinco = "0" + fecha_cinco;
                        }
                        fecha_uno = año + "-" + mes + "-" + fecha_cinco;
                        fecha_dos = fecha_cinco + "/" + mes + "/" + año;

                        conexionmy.Open();
                        SqlCommand comando_crea = new SqlCommand("USE " + server +
                        " truncate table prestamos_poliza " +
                        " ", conexionmy);
                        comando_crea.ExecuteNonQuery();
                        conexionmy.Close();
                        //MessageBox.Show("" + fecha_cinco);
                        conexionmy.Open();
                        SqlCommand simbolo = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + letra + "')", conexionmy);
                        simbolo.ExecuteNonQuery();
                        conexionmy.Close();

                        conexionmy.Open();
                        SqlCommand diadepoliza = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + fecha_cinco + "')", conexionmy);
                        diadepoliza.ExecuteNonQuery();
                        conexionmy.Close();

                        conexionmy.Open();
                        SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "COBROS DEL  " + fecha_dos + " " + NUMERO_DE_CAJA + "')", conexionmy);
                        leyendaP.ExecuteNonQuery();
                        conexionmy.Close();

                        conexionmy.Open();
                        SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + LOCALIDAD + " " + "CAJA # " + NUMERO_DE_CAJA + "')", conexionmy);
                        leyendaPL.ExecuteNonQuery();
                        conexionmy.Close();

                        poliza_cobro();
                    }//for 

                
             

            }//try
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void poliza_cobro()
        {


            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            /*conexionmy.ConnectionString = sqlcnx;//conexion mysql*/


            fecha_tres = año;
            fecha_cuatro = mes;

            fecha_dos = fecha_cinco + "/" + mes + "/" + año;
            int año_inicial, nom, valor;
            string seis = "";
            string cuatro;
            string catorce;
            año_inicial = 2006;
            nom = 1;
            valor = Convert.ToInt32(año);
            cuatro = "0" + mes;
            while (año_inicial < valor)
            {
                nom = nom + 1;
                año_inicial = año_inicial + 1;

            }
            if (nom < 10)
            {
                seis = Convert.ToString("00" + nom);
            }
            if (nom >= 10)
            {
                seis = Convert.ToString("0" + nom);
            }
            catorce = CUENTA_PRESTAMOS + "-" + seis + "-" + cuatro; //cuanta para prestamos
            string letra = "DESEMP";
            DataTable tablacobros = new DataTable();
            tablacobros.Clear();
            SqlDataAdapter comando_carga = new SqlDataAdapter("USE " + server +
           " SELECT facturas.factura, facturas.contrato, " + caja + ".debe, facturas.importefact, facturas.ivafact, facturas.totalfact, facturas.status FROM  facturas " +
                " INNER JOIN " + caja + " ON facturas.factura = " + caja + ".folio WHERE " +
                " facturas.FechaFact='" + Convert.ToDateTime(fecha_uno).ToString("yyyy-dd-MM") + "' and " + caja + ".concepto like '%" + letra + "%' order by facturas.factura asc" +
           "", conexionmy);

            //comienza el if de lleno o vacio
            comando_carga.Fill(tablacobros);
            //this.dataGridView1.DataSource = tablacobros;

            //veo si esta vacio el dia
            if (tablacobros.Rows.Count == 0)
            {
                //exportar con nota de no hubo operacion este dia y seguir con rutina
            }
            else
            {
                string FACTURA, CONTRATO, DEBE, INTERES, IVA, TOTALFACT, STATUS, FECHACONS, PRESTAMO = "0", quince = "0";
                foreach (DataRow dr in tablacobros.Rows)
                {


                    FACTURA = dr[0].ToString();//contrato
                    CONTRATO = dr[1].ToString(); //status
                    DEBE = dr[2].ToString(); //prestamo $$
                    INTERES = dr[3].ToString();//contrato
                    IVA = dr[4].ToString(); //status
                    TOTALFACT = dr[5].ToString(); //prestamo $$
                    STATUS = dr[6].ToString();//contratodr.Cells[6].Value.ToString();//contrato

                    //NUMERO DE LA CUENTA
                    conexionmy.Open();
                    SqlCommand comando_A = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + CUENTA_CAJA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    comando_A.ExecuteNonQuery();
                    conexionmy.Close();
                    //LAS LEYENDAS
                    conexionmy.Open();
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + FACTURA + "   " + "C-" + CONTRATO + "')", conexionmy);
                    comando_B.ExecuteNonQuery();
                    conexionmy.Close();
                    //EL MONTO DEL PRESTAMO
                    string monto_valor;
                    //monto_valor = Convert.ToString(DEBE);
                    //monto_valor = string.Format(monto_valor, "####0.000000");
                    decimal unosidebed = decimal.Round(Convert.ToDecimal(DEBE), 2, MidpointRounding.AwayFromZero);
                    monto_valor = string.Format("{0:0.000000}", unosidebed);


                    if (STATUS == "CANCELADO")
                    {
                        conexionmy.Open();
                        SqlCommand comando_C = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        comando_C.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_D = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + monto_valor + ",1.00" + "')", conexionmy);
                        comando_D.ExecuteNonQuery();
                        conexionmy.Close();
                    }

                    //ahora busco le fecha de origen del contrato en cuestion

                    DataTable fecha_con = new DataTable();
                    fecha_con.Clear();
                    SqlDataAdapter fecha_Carga = new SqlDataAdapter("USE " + server +
                   " SELECT fechacons, prestamo FROM  contratos where contrato='" + CONTRATO + "' " +
                   "", conexionmy);
                    fecha_Carga.Fill(fecha_con);
                   // this.dataGridView2.DataSource = fecha_con;
                    if (fecha_con.Rows.Count == 0)
                    {
                        //exportar con nota de no hubo operacion este dia y seguir con rutina
                    }
                    else
                    {
                        FECHACONS = fecha_con.Rows[0][0].ToString(); //this.dataGridView2.Rows[0].Cells[0].Value.ToString();//fechacons
                        PRESTAMO = fecha_con.Rows[0][1].ToString();//this.dataGridView2.Rows[0].Cells[1].Value.ToString();//prestamo

                        int año_inicial_uno, nom_uno, valor_uno;
                        string seis_uno = "";
                        string cuatro_uno;

                        string año_fechacons, mes_fechacons;
                        año_fechacons = Convert.ToDateTime(FECHACONS).ToString("yyyy");
                        mes_fechacons = Convert.ToDateTime(FECHACONS).ToString("MM");


                        año_inicial_uno = 2006;
                        nom_uno = 1;
                        valor_uno = Convert.ToInt32(año_fechacons);//año actual

                        cuatro_uno = "0" + mes_fechacons;

                        while (año_inicial_uno < valor_uno)
                        {
                            nom_uno = nom_uno + 1;
                            año_inicial_uno = año_inicial_uno + 1;

                        }
                        if (nom_uno < 10)
                        {
                            seis_uno = Convert.ToString("00" + nom_uno);
                        }
                        if (nom_uno >= 10)
                        {
                            seis_uno = Convert.ToString("0" + nom_uno);
                        }

                        quince = CUENTA_PRESTAMOS + "-" + seis_uno + "-" + cuatro_uno;

                    }

                    //

                    //AHORA LA CUENTA DE LA CAJA
                    conexionmy.Open();
                    SqlCommand comando_E = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + quince + "                ," + "  " + DEPARTAMENTO + " ')", conexionmy);
                    comando_E.ExecuteNonQuery();
                    conexionmy.Close();
                    //ahora leyenda caja
                    conexionmy.Open();
                    SqlCommand comando_F = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + FACTURA + "   " + "C-" + CONTRATO + "')", conexionmy);
                    comando_F.ExecuteNonQuery();
                    conexionmy.Close();
                    //inserto espacio
                    //ahora la el monto de prestamo
                    conexionmy.Open();
                    SqlCommand comando_G = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('')", conexionmy);
                    comando_G.ExecuteNonQuery();
                    conexionmy.Close();
                    //nuevamene la cantidad
                    //ahora la el monto de prestamo             
                    //string h1;
                    string valor_convertido;
                    //h1 = Convert.ToString(PRESTAMO);
                    //valor_convertido = string.Format(h1, "####0.000000");

                    decimal unosidebede = decimal.Round(Convert.ToDecimal(PRESTAMO), 2, MidpointRounding.AwayFromZero);
                    valor_convertido = string.Format("{0:0.000000}", unosidebede);

                    if (STATUS == "CANCELADO")
                    {
                        conexionmy.Open();
                        SqlCommand comando_H = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        comando_H.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_I = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + valor_convertido + ",1.00" + "')", conexionmy);
                        comando_I.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    //'ahora cuenta caja
                    conexionmy.Open();
                    SqlCommand comando_Ia = new SqlCommand("USE " + server + " INSERT into prestamos_poliza(cuenta) VALUES('" + CUENTA_INTERES + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    comando_Ia.ExecuteNonQuery();
                    conexionmy.Close();

                    //ahora leyenda interes
                    conexionmy.Open();
                    SqlCommand comando_Iaa = new SqlCommand("USE " + server + " INSERT into prestamos_poliza(cuenta) VALUES('" + "INTERESES- " + FACTURA + " " + "C-" + CONTRATO + "')", conexionmy);
                    comando_Iaa.ExecuteNonQuery();
                    conexionmy.Close();


                    //un nuevo espacio
                    conexionmy.Open();
                    SqlCommand comando_Iaabk = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta) VALUES(' ')", conexionmy);
                    comando_Iaabk.ExecuteNonQuery();
                    conexionmy.Close();
                    //ingreso del interes    
                    //string h13;
                    string valor_convertido3;
                    //h13 = Convert.ToString(INTERES);
                    //valor_convertido3 = string.Format(h13, "####0.000000");

                    decimal unosidebedep = decimal.Round(Convert.ToDecimal(INTERES), 2, MidpointRounding.AwayFromZero);
                    valor_convertido3 = string.Format("{0:0.000000}", unosidebedep);

                    if (STATUS == "CANCELADO")
                    {
                        conexionmy.Open();
                        SqlCommand comando_Ha = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        comando_Ha.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_Iz = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + valor_convertido3 + ",1.00" + "')", conexionmy);
                        comando_Iz.ExecuteNonQuery();
                        conexionmy.Close();
                    }

                    //'ahora la cuenta del iva
                    conexionmy.Open();
                    SqlCommand comando_Iza = new SqlCommand("USE " + server + " INSERT into prestamos_poliza(cuenta) VALUES('" + CUENTA_IVA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    comando_Iza.ExecuteNonQuery();
                    conexionmy.Close();
                    //leyenda iva
                    conexionmy.Open();
                    SqlCommand comando_Iaabe = new SqlCommand("USE " + server + " INSERT into prestamos_poliza(cuenta) VALUES('" + "IVA- " + FACTURA + "   " + "C-" + CONTRATO + "')", conexionmy);
                    comando_Iaabe.ExecuteNonQuery();
                    conexionmy.Close();
                    //un nuevo espacio
                    conexionmy.Open();
                    SqlCommand comando_Iaaba = new SqlCommand("USE " + server + " INSERT into prestamos_poliza(cuenta)VALUES(' ')", conexionmy);
                    comando_Iaaba.ExecuteNonQuery();
                    conexionmy.Close();

                    //'valor del iva
                    //string h135;
                    string valor_convertido35;
                    //h135 = Convert.ToString(IVA);
                    //valor_convertido35 = string.Format(h135, "####0.000000");

                    decimal unosidebedepa = decimal.Round(Convert.ToDecimal(IVA), 2, MidpointRounding.AwayFromZero);
                    valor_convertido35 = string.Format("{0:0.000000}", unosidebedepa);

                    if (STATUS == "CANCELADO")
                    {
                        conexionmy.Open();
                        SqlCommand comando_Ha = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        comando_Ha.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_Iz = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + valor_convertido35 + ",1.00" + "')", conexionmy);
                        comando_Iz.ExecuteNonQuery();
                        conexionmy.Close();
                    }

                    //
                }//cierre del ciclo for each
                conexionmy.Open();
                SqlCommand comando_J = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('FIN')", conexionmy);
                comando_J.ExecuteNonQuery();
                conexionmy.Close();
                notas_cobro();
            }//if del rows
        }

        public void notas_cobro()
        {
            SqlConnection conexionmy = new SqlConnection();
            conexionmy.ConnectionString = sqlcnx;//conexion mysql
            DataTable tablaprestamos = new DataTable();
            tablaprestamos.Clear();
            SqlDataAdapter datos_presa = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  prestamos_poliza " +
                " order by  no asc " +
                "", conexionmy);
            datos_presa.Fill(tablaprestamos);
            //this.dataGridView2.DataSource = tablaprestamos;
            System.IO.StreamWriter escribir = new System.IO.StreamWriter(path + "/NOTADEPAGO " + caja + "-" + fecha_uno + ".pol");
            foreach (DataRow fila in tablaprestamos.Rows)
                escribir.WriteLine(fila[1]);
            escribir.Close();

        }
        #endregion


        #region Interes Diario
        public void diario_interes()
        {
            try
            {
                //if (checkBoxX3.Checked == true)
                //{
                    //2. cargo las cajas de la sucursal seleccionada
                    SqlConnection conexionmy = new SqlConnection(sqlcnx);
                // conexionmy.ConnectionString = sqlcnx;//conexion mysql

                progressBarX3.Maximum = dias_en_mes;
                progressBarX3.Value = 0;
                //string CONTRATO, STATUS, PRESTAMO;
                for (int nveces = 1; nveces <= dias_en_mes; nveces++)
                    {

                    backgroundWorker4.ReportProgress(nveces);

                    fecha_cinco = Convert.ToString(nveces);
                        if (fecha_cinco.Length == 1)
                        {
                            fecha_cinco = "0" + fecha_cinco;
                        }
                        fecha_uno = año + "-" + mes + "-" + fecha_cinco;
                        fecha_dos = fecha_cinco + "/" + mes + "/" + año;

                        conexionmy.Open();
                        SqlCommand comando_crea = new SqlCommand("USE " + server +
                        " truncate table prestamos_poliza " +
                        " ", conexionmy);
                        comando_crea.ExecuteNonQuery();
                        conexionmy.Close();
                        //MessageBox.Show("" + fecha_cinco);
                        conexionmy.Open();
                        SqlCommand simbolo = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + letra + "')", conexionmy);//letra
                        simbolo.ExecuteNonQuery();
                        conexionmy.Close();

                        conexionmy.Open();
                        SqlCommand diadepoliza = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + fecha_cinco + "')", conexionmy);//dia
                        diadepoliza.ExecuteNonQuery();
                        conexionmy.Close();
                        ////este es diferente con iff segn la empresa
                        //LAS LEYENDAS con if segun empr
                        if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                        {
                            conexionmy.Open();
                            SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Interes Diario   " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + "')", conexionmy);
                            leyendaP.ExecuteNonQuery();
                            conexionmy.Close();

                        }
                        else if (empresa_Conta == "MONTE ROS SA DE CV")
                        {
                            conexionmy.Open();
                            SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Interes Diario   " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + "')", conexionmy);
                            leyendaP.ExecuteNonQuery();
                            conexionmy.Close();

                        }
                        else
                        {
                            conexionmy.Open();
                            SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Interes Diario   " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + "')", conexionmy);
                            leyendaP.ExecuteNonQuery();
                            conexionmy.Close();

                        }

                        conexionmy.Open();
                        SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('.')", conexionmy);
                        leyendaPL.ExecuteNonQuery();
                        conexionmy.Close();

                        poliza_diario();
                    }//for 

                    //MessageBox.Show("Polizas de diario realizadas", "Polizas SEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //}//if
                //else
                //{
                //    semanal_interes();
                //}

            }//try
            catch (Exception ex)
            {
                 MessageBox.Show(ex.Message);
            }
        }

        public void poliza_diario()
        {
            //2. cargo las cajas de la sucursal seleccionada
            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            //conexionmy.ConnectionString = sqlcnx;//conexion mysql

            switch (mes)
            {
                case "01":
                    mes_letra = "Enero";
                    break;
                case "02":
                    mes_letra = "Febrero";
                    break;
                case "03":
                    mes_letra = "Marzo";
                    break;
                case "04":
                    mes_letra = "Abril";
                    break;
                case "05":
                    mes_letra = "Mayo";
                    break;
                case "06":
                    mes_letra = "Junio";
                    break;
                case "07":
                    mes_letra = "Julio";
                    break;
                case "08":
                    mes_letra = "Agosto";
                    break;
                case "09":
                    mes_letra = "Septiembre";
                    break;
                case "10":
                    mes_letra = "Octubre";
                    break;
                case "11":
                    mes_letra = "Noviembre";
                    break;
                case "12":
                    mes_letra = "Diciembre";
                    break;
                default:
                    break;

            }

            //string letra = "DESEMP";
            DataTable tablacobros = new DataTable();
            tablacobros.Clear();
            SqlDataAdapter comando_carga = new SqlDataAdapter("USE " +server+
           "  SELECT sum(ImporteFact), sum(IVAFACT) FROM  facturas INNER JOIN " + caja + " ON facturas.factura = " + caja + ".folio " +
            " AND facturas.STATUS = 'VALIDO'  WHERE facturas.FechaFact= '" + Convert.ToDateTime(fecha_uno).ToString("yyyy-dd-MM") + "' " +
            " and " + caja + ".concepto like '%DESEMP%'" +
           "", conexionmy);
            //comienza el if de lleno o vacio
            comando_carga.Fill(tablacobros);
            //this.dataGridView1.DataSource = tablacobros;
            //veo si esta vacio el dia
            if (tablacobros.Rows.Count == 0)
            {

                //exportar con nota de no hubo operacion este dia y seguir con rutina
            }
            else
            {

                string importe_del_iva, iva_solo;



                importe_del_iva = tablacobros.Rows[0][0].ToString();// tablacobros.Rows[0].ItemArray[0].ToString();//contrato
                iva_solo =tablacobros.Rows[0][1].ToString();// tablacobros.Rows[0].ItemArray[1].ToString();//status

                if (importe_del_iva == "")
                {
                    importe_del_iva = "0.00";
                    iva_solo = "0.00";


                }


                //NUMERO DE LA CUENTA
                conexionmy.Open();
                SqlCommand comando_A = new SqlCommand("USE " + server + " INSERT into prestamos_poliza(cuenta) VALUES('" + CUENTA_INTERES + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                comando_A.ExecuteNonQuery();
                conexionmy.Close();


                //LAS LEYENDAS con if segun empr
                if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + "')", conexionmy);
                    comando_B.ExecuteNonQuery();
                    conexionmy.Close();
                }
                else if (empresa_Conta == "MONTE ROS SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_B.ExecuteNonQuery();
                    conexionmy.Close();
                }
                else
                {
                    conexionmy.Open();
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_B.ExecuteNonQuery();
                    conexionmy.Close();
                }

                //EL MONTO DEL PRESTAMO me quede donde dice monteo del importe en notas d diarii linea 3236 en vb net
                string monto_valor_T;
                //monto_valor_T = Convert.ToString(string.Format("{0:F2}", importe_del_iva)); //Convert.ToString(importe_del_iva);
                //monto_valor_T = string.Format(monto_valor_T, "####0.000000");

                decimal unosideb = decimal.Round(Convert.ToDecimal(importe_del_iva), 2, MidpointRounding.AwayFromZero);
                monto_valor_T = string.Format("{0:0.000000}", unosideb);


                conexionmy.Open();
                SqlCommand comando_D = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + monto_valor_T + ",1.00" + "')", conexionmy);
                comando_D.ExecuteNonQuery();
                conexionmy.Close();

                conexionmy.Open();
                SqlCommand comando_D_D = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + CUENTA_INTERES + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                comando_D_D.ExecuteNonQuery();
                conexionmy.Close();

                //LAS LEYENDAS2 segun la empresa entr otr if
                if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_B_b = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + "')", conexionmy);
                    comando_B_b.ExecuteNonQuery();
                    conexionmy.Close();
                }
                else if (empresa_Conta == "MONTE ROS SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_B_b = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_B_b.ExecuteNonQuery();
                    conexionmy.Close();
                }
                else
                {
                    conexionmy.Open();
                    SqlCommand comando_B_b = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_B_b.ExecuteNonQuery();
                    conexionmy.Close();
                }


                conexionmy.Open();
                SqlCommand comando_B_e = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES(' ')", conexionmy);
                comando_B_e.ExecuteNonQuery();
                conexionmy.Close();

                string monto_valor_G;
                // monto_valor_G = Convert.ToString(string.Format("{0:F2}", importe_del_iva)); //Convert.ToString(importe_del_iva);
                //monto_valor_G = string.Format(monto_valor_G, "####0.000000");

                decimal unosidebe = decimal.Round(Convert.ToDecimal(importe_del_iva), 2, MidpointRounding.AwayFromZero);
                monto_valor_G = string.Format("{0:0.000000}", unosidebe);

                conexionmy.Open();
                SqlCommand comando_D_H = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + monto_valor_G + ",1.00" + "')", conexionmy);
                comando_D_H.ExecuteNonQuery();
                conexionmy.Close();

                conexionmy.Open();
                SqlCommand comando_D_R = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + CUENTA_IVA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                comando_D_R.ExecuteNonQuery();
                conexionmy.Close();
                //if segun la empresa
                //LAS LEYENDAS2 segun la empresa entr otr if
                if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_D_S = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + "')", conexionmy);
                    comando_D_S.ExecuteNonQuery();
                    conexionmy.Close();

                }
                else if (empresa_Conta == "MONTE ROS SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_D_S = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_D_S.ExecuteNonQuery();
                    conexionmy.Close();

                }
                else
                {
                    conexionmy.Open();
                    SqlCommand comando_D_S = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_D_S.ExecuteNonQuery();
                    conexionmy.Close();

                }

                //conexionmy.Open();
                //SqlCommand comando_D_S_e = new SqlCommand("USE " + this.label4.Text + " INSERT INTO prestamos_poliza(cuenta)VALUES(' ')", conexionmy);
                //comando_D_S_e.ExecuteNonQuery();
                //conexionmy.Close();

                string monto_valor_GG;
                //monto_valor_GG = Convert.ToString(string.Format("{0:F2}", iva_solo)); //Convert.ToString(iva_solo);
                //monto_valor_GG = string.Format(monto_valor_GG, "####0.000000");
                decimal unosidebed = decimal.Round(Convert.ToDecimal(iva_solo), 2, MidpointRounding.AwayFromZero);
                monto_valor_GG = string.Format("{0:0.000000}", unosidebed);


                conexionmy.Open();
                SqlCommand comando_D_S_R = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + monto_valor_GG + ",1.00" + "')", conexionmy);
                comando_D_S_R.ExecuteNonQuery();
                conexionmy.Close();

                conexionmy.Open();
                SqlCommand comando_D_S_Ru = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + CUENTA_IVA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                comando_D_S_Ru.ExecuteNonQuery();
                conexionmy.Close();


                //LAS LEYENDAS2 segun la empresa entr otr if
                if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_D_Sw = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + "')", conexionmy);
                    comando_D_Sw.ExecuteNonQuery();
                    conexionmy.Close();

                }
                else if (empresa_Conta == "MONTE ROS SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_D_Sw = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_D_Sw.ExecuteNonQuery();
                    conexionmy.Close();

                }
                else
                {
                    conexionmy.Open();
                    SqlCommand comando_D_Sw = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_D_Sw.ExecuteNonQuery();
                    conexionmy.Close();

                }
                conexionmy.Open();
                SqlCommand comando_D_S_F = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES(' ')", conexionmy);
                comando_D_S_F.ExecuteNonQuery();
                conexionmy.Close();

                string monto_valor_GGe;
                //monto_valor_GGe = Convert.ToString(string.Format("{0:F2}", iva_solo));// Convert.ToString(iva_solo);
                //monto_valor_GGe = string.Format(monto_valor_GGe, "####0.000000");
                decimal unosidebede = decimal.Round(Convert.ToDecimal(iva_solo), 2, MidpointRounding.AwayFromZero);
                monto_valor_GGe = string.Format("{0:0.000000}", unosidebede);

                conexionmy.Open();
                SqlCommand comando_D_S_Ra = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + monto_valor_GGe + ",1.00" + "')", conexionmy);
                comando_D_S_Ra.ExecuteNonQuery();
                conexionmy.Close();

                conexionmy.Open();
                SqlCommand comando_D_S_Fi = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('FIN')", conexionmy);
                comando_D_S_Fi.ExecuteNonQuery();
                conexionmy.Close();
                notas_diario();
            }

        }
        public void notas_diario()
        {
            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            //conexionmy.ConnectionString = sqlcnx;//conexion mysql
            DataTable tablaprestamos = new DataTable();
            tablaprestamos.Clear();
            SqlDataAdapter datos_presa = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  prestamos_poliza " +
                " order by  no asc " +
                "", conexionmy);
            datos_presa.Fill(tablaprestamos);
            //this.dataGridView2.DataSource = tablaprestamos;
            System.IO.StreamWriter escribir = new System.IO.StreamWriter(path + "/DIARIO " + caja + "-" + fecha_uno + ".pol");
            foreach (DataRow fila in tablaprestamos.Rows)
                escribir.WriteLine(fila[1]);
            escribir.Close();
        }


       
        #endregion





        void CheckedDesCheckedHabilitarDesHabilitar()
        {

            //Todos los checkbox serán considerados como un item de checkboxlist:
            List<CheckBoxX> chkLst = new List<CheckBoxX>();

            //Añadimos todos los checkboxes al checkboxlist:
            chkLst.AddRange(this.groupPanel1.Controls.OfType<CheckBoxX>());

            for (int i = 0; i < chkLst.Count; i++)
            {

                //Hacemos el Deschecked a todos los checkbox:
                chkLst[i].Checked = true;

                //Si quisieramos que todos los chekbox estuvieran en checked:
                //chkLst[i].Checked = true

                //Si quisieramos que todos los chekbox se habilitaran:
                //chkLst[i].Enabled = true

                //Si quisieramos que todos los chekbox se deshabilitaran:
                //chkLst[i].Enabled = false


            }

        }

        private void backgroundWorker2_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
           
            prestamos();
               
        }

        private void backgroundWorker2_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBarX1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            nveces = 0;
            MessageBox.Show("Poliza Prestamos Terminada", "EKPolizasSemp", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void backgroundWorker3_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            nveces = 0;
            MessageBox.Show("Poliza Cobros Terminada", "EKPolizasSemp", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void backgroundWorker3_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBarX2.Value = e.ProgressPercentage;
        }

        private void backgroundWorker3_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            cobros();
        }

        private void backgroundWorker4_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            nveces = 0;
            MessageBox.Show("Poliza Interes Diario Terminada", "EKPolizasSemp", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void backgroundWorker4_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBarX3.Value = e.ProgressPercentage;
        }

        private void backgroundWorker4_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            diario_interes();
        }

        private void comboBoxEx4_SelectedIndexChanged(object sender, EventArgs e)
        {
          
        }

    

        private void buttonX1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialogoRuta = new FolderBrowserDialog();

            if (dialogoRuta.ShowDialog() == DialogResult.OK)
            {

                path = dialogoRuta.SelectedPath;
                textBoxX1.Text = path;

            }
        }

        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            paso2();
        }

        #endregion

      

    }
}
