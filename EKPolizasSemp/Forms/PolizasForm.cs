

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

    public partial class PolizasForm : Office2007Form
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
        string _mesLetra = "";
        int nveces = 0;
        int nveces2 = 0;
        int nveces3 = 0;

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
        string letra2;
        string letra3;
        string letra4;
        string letra5;
        string letra6;
        string letra7;

        DataTable tablecaja_dos = new DataTable();

        int activarCombo = 0;

        #endregion

        #region Events (eventos)
        private void Form1_Load(object sender, EventArgs e)
        {

            circularProgress1.Visible = true;
            circularProgress1.IsRunning = true;
            CheckForIllegalCrossThreadCalls = false;
            backgroundWorker1.RunWorkerAsync();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            ExisteTabla(server, sqlcnx);
            FolderBrowserDialog dialogoRuta = new FolderBrowserDialog();


            if (dialogoRuta.ShowDialog() == DialogResult.OK)
            {

                path = dialogoRuta.SelectedPath;
                textBoxX1.Text = path;

            }
        }

        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (activarCombo == 0)
            {
                activarCombo = 1;

            }
            else
            {
                paso2();
            }

        }
        //GENERATE
        private void buttonX2_Click(object sender, EventArgs e)
        {
            caja = this.comboBoxEx4.Text;
            ExisteTabla(server, sqlcnx);
            if (!string.IsNullOrEmpty(textBoxX1.Text) && !string.IsNullOrEmpty(txtPorcentaje.Text))
            {
                progressBarX1.Value = 0;
                progressBarX2.Value = 0;
                progressBarX3.Value = 0;
                progressBarX4.Value = 0;
                progressBarX5.Value = 0;
                progressBarX6.Value = 0;
                progressBarX7.Value = 0;



                mes_calculo();

                progressBarX1.Maximum = dias_en_mes;
                progressBarX2.Maximum = dias_en_mes;
                progressBarX3.Maximum = dias_en_mes;
                progressBarX4.Maximum = dias_en_mes;
                progressBarX5.Maximum = dias_en_mes;
                progressBarX6.Maximum = dias_en_mes;
                progressBarX7.Maximum = dias_en_mes;

                DialogResult pregunta = MessageBox.Show("Generar Polizas?",
                                                        "EKPolizasSemp",
                                                        MessageBoxButtons.YesNo,
                                                        MessageBoxIcon.Question);
                if (pregunta == DialogResult.Yes)
                {


                    // caja = this.comboBoxEx4.Text;
                    server = this.label4.Text;
                    path = this.textBoxX1.Text;

                    crear_poliza();
                }

            }
            else
            {
                MessageBox.Show("Ingresa una ruta valida para guardar polizas y Porcentaje para Remisiones",
                                                      "EKPolizasSemp",
                                                      MessageBoxButtons.OK,
                                                      MessageBoxIcon.Information);
                return;
            }



        }
        //SELECT BOX
        private void comboBoxEx4_SelectedIndexChanged(object sender, EventArgs e)
        {
            // caja = this.comboBoxEx4.Text;
        }



        #endregion

        #region Methods (metodos)



        #region Metodos Generales
        public PolizasForm()
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
            backgroundWorker5.WorkerReportsProgress = true;
            backgroundWorker5.WorkerSupportsCancellation = true;
            backgroundWorker6.WorkerReportsProgress = true;
            backgroundWorker6.WorkerSupportsCancellation = true;
            backgroundWorker7.WorkerReportsProgress = true;
            backgroundWorker7.WorkerSupportsCancellation = true;
            backgroundWorker8.WorkerReportsProgress = true;
            backgroundWorker8.WorkerSupportsCancellation = true;
        }

        private void crear_poliza()
        {


            if (checkBoxX1.Checked == true)
            {
                poliza(1);

            }

            else if (checkBoxX2.Checked == true)
            {
                poliza(2);

            }

            else if (checkBoxX3.Checked == true)
            {
                poliza(3);

            }

            else if (checkBoxX4.Checked == true)
            {
                poliza(4);

            }

            else if (checkBoxX5.Checked == true)
            {
                poliza(5);

            }

            else if (checkBoxX6.Checked == true)
            {
                poliza(6);

            }

            else if (checkBoxX7.Checked == true)
            {
                poliza(7);

            }

            else
            {
                MessageBox.Show("Polizas Terminadas",
                                                    "EKPolizasSemp",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Information);
            }


            #region trash
            ////Todos los checkbox serán considerados como un item de checkboxlist:
            //List<CheckBoxX> chkLst = new List<CheckBoxX>();

            ////Añadimos todos los checkboxes al checkboxlist:
            //chkLst.AddRange(this.groupPanel1.Controls.OfType<CheckBoxX>());

            //for (int i = 0; i < chkLst.Count; i++)
            //{

            //    if(chkLst[i].Checked==true && chkLst[i].Text == "Prestamos")
            //    {
            //        opcion = 1;
            //        letra = this.textBoxX3.Text;
            //        backgroundWorker2.RunWorkerAsync();
            //    }

            //    if (chkLst[i].Checked == true && chkLst[i].Text == "Cobros")
            //    {
            //        opcion = 2;
            //        letra = this.textBoxX2.Text;
            //        backgroundWorker3.RunWorkerAsync();
            //    }

            //    if (chkLst[i].Checked == true && chkLst[i].Text == "Diario")
            //    {
            //        opcion = 3;
            //        letra = this.textBoxX4.Text;
            //        backgroundWorker4.RunWorkerAsync();
            //    }


            //    if (chkLst[i].Checked == true && chkLst[i].Text == "Interes Semanal")
            //    {

            //    }

            //    if (chkLst[i].Checked == true && chkLst[i].Text == "Remision Semanal")
            //    {

            //    }


            //Hacemos el Deschecked a todos los checkbox:
            // chkLst[i].Checked = true;

            //Si quisieramos que todos los chekbox estuvieran en checked:
            //chkLst[i].Checked = true

            //Si quisieramos que todos los chekbox se habilitaran:
            //chkLst[i].Enabled = true

            //Si quisieramos que todos los chekbox se deshabilitaran:
            //chkLst[i].Enabled = false


            //} 
            #endregion

        }
        private void poliza(int valor)
        {

            switch (valor)
            {
                case 1://prestamos


                    letra = textBoxX3.Text;
                    backgroundWorker2.RunWorkerAsync();
                    break;
                case 2:


                    letra2 = textBoxX2.Text;
                    backgroundWorker3.RunWorkerAsync();
                    break;
                case 3:
                    letra3 = textBoxX4.Text;
                    backgroundWorker4.RunWorkerAsync();
                    break;
                case 4:
                    letra4 = textBoxX5.Text;
                    backgroundWorker5.RunWorkerAsync();

                    break;
                case 5:
                    letra5 = textBoxX6.Text;
                    backgroundWorker6.RunWorkerAsync();
                    break;

                case 6:
                    letra6 = textBoxX7.Text;
                    backgroundWorker7.RunWorkerAsync();
                    break;

                case 7:
                    letra7 = textBoxX8.Text;
                    backgroundWorker8.RunWorkerAsync();
                    break;

                default:
                    break;
            }
        }
        #endregion

        #region LLenar Formulario
        private void ExisteTabla(string server, string conexion)
        {

            SqlConnection conexionmy = new SqlConnection(conexion);

            conexionmy.Open();
            SqlCommand simbolo = new SqlCommand("USE " + server + " truncate table cobros_poliza ", conexionmy);
            simbolo.ExecuteNonQuery();
            conexionmy.Close();


         

            conexionmy.Open();
            SqlCommand simbolo2 = new SqlCommand("USE " + server + " truncate table diario_poliza", conexionmy);
            simbolo2.ExecuteNonQuery();
            conexionmy.Close();


         

            conexionmy.Open();
            SqlCommand simbolo3 = new SqlCommand("USE " + server + " Truncate table interesSemanal_poliza ", conexionmy);
            simbolo3.ExecuteNonQuery();
            conexionmy.Close();

            conexionmy.Open();
            SqlCommand simbolo4 = new SqlCommand("USE " + server + "  Truncate Table remisionSemanal_poliza ", conexionmy);
            simbolo4.ExecuteNonQuery();
            conexionmy.Close();

         

            conexionmy.Open();
            SqlCommand simbolo5 = new SqlCommand("USE " + server + " truncate Table  cobros_poliza_interes ", conexionmy);
            simbolo5.ExecuteNonQuery();
            conexionmy.Close();

            conexionmy.Open();
            SqlCommand simbolo6 = new SqlCommand("USE " + server + " truncate Table  cobros_poliza_desempenos ", conexionmy);
            simbolo6.ExecuteNonQuery();
            conexionmy.Close();


        }

        private void paso2()
        {

            //1. busco la tabla de datos en el servidor segun la sucursal seleccionada

            try
            {
                SqlConnection conexionmy = new SqlConnection(sqlcnx + ";Connection Timeout=30");
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
                tablecaja_dos.Clear();

                SqlDataAdapter datoscaja = new SqlDataAdapter("USE " + this.label4.Text + " " +
                    " Select NumCaja from selcaja order by NumCaja asc", conexionmy);
                datoscaja.Fill(tablecaja_dos);
                comboBoxEx4.ValueMember = "NumCaja";
                comboBoxEx4.DisplayMember = "NumCaja";
                comboBoxEx4.DataSource = tablecaja_dos;

                server = label4.Text;

                //this.dataGridView5.DataSource = tablecaja_dos;

                //ExisteTabla();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString() + "-No encontre datos de localidad intenta con una diferente por favor");
            }



        }

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
            año = new String[] { "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026", "2027", "2028", "2029", "2030" };
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

        private void conexion()
        {
            //aligual que las demas aplicaciones cargaremos nuestra llave al servidor de oficinas para la conexion directa
            //string cadena = "C:/SEMP2013/GeneradorPolizasSEMP/Polizas_General/Polizas_General/bin/Debug/cdblista.txt";
            string cadena = "C:/SEMP2013/cdb/cdb.txt";
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
                        MessageBox.Show("Error de Conexion", "Polizas Semp", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }






            }

        }

        private void mes_calculo()
        {
            //1obtener la fecha inicial
            caja = comboBoxEx4.Text;
            switch (this.comboBoxEx2.Text)
            {
                case "ENERO":
                    mes = "01";
                    _mesLetra = "ene.";
                    break;
                case "FEBRERO":
                    mes = "02";
                    _mesLetra = "feb.";
                    break;
                case "MARZO":
                    mes = "03";
                    _mesLetra = "mar.";
                    break;
                case "ABRIL":
                    mes = "04";
                    _mesLetra = "abr.";
                    break;
                case "MAYO":
                    mes = "05";
                    _mesLetra = "may.";
                    break;
                case "JUNIO":
                    mes = "06";
                    _mesLetra = "jun.";
                    break;
                case "JULIO":
                    mes = "07";
                    _mesLetra = "jul.";
                    break;
                case "AGOSTO":
                    mes = "08";
                    _mesLetra = "ago.";
                    break;
                case "SEPTIEMBRE":
                    mes = "09";
                    _mesLetra = "sep.";
                    break;
                case "OCTUBRE":
                    mes = "10";
                    _mesLetra = "oct.";
                    break;
                case "NOVIEMBRE":
                    mes = "11";
                    _mesLetra = "nov.";
                    break;
                case "DICIEMBRE":
                    mes = "12";
                    _mesLetra = "dic.";
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
            SqlDataAdapter datosSql = new SqlDataAdapter("USE SEMP2013_NAU_1  " +
                "Select *  from contabilidad where leyenda2='" + caja + "'" +
                "", conexionmy);

            //SqlDataAdapter datosSql = new SqlDataAdapter("USE " + server + "  " +
            //    "Select *  from contabilidad where leyenda2='" + caja + "'" +
            //    "", conexionmy);
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

        #region Prestamos rutina
        //INICIO DE GENERACION DE ARCHIVO INTERES
        private void prestamos()
        {



            try
            {

                letra = textBoxX3.Text;
                //progressBarX1.Maximum = dias_en_mes;
                //progressBarX1.Value = 0;
                int _nveces;
                _nveces = 1;
                if (!string.IsNullOrEmpty(txtDias.Text) && !string.IsNullOrEmpty(txtDiasFinal.Text))
                {

                    if (int.Parse(txtDias.Text) > 0 && int.Parse(txtDiasFinal.Text) > 0)
                    {
                        _nveces = int.Parse(txtDias.Text);
                        dias_en_mes = int.Parse(txtDiasFinal.Text);
                    }
                    else
                    {
                        _nveces = 1;
                    }

                }




                for (nveces = _nveces; nveces <= dias_en_mes; nveces++)
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

                    fecha_dos = fecha_cinco + "/" + _mesLetra + "/" + año;

                    conexionmy.Open();
                    SqlCommand simbolo = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + letra + "')", conexionmy);
                    simbolo.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand diadepoliza = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + fecha_cinco + "')", conexionmy);
                    diadepoliza.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    //SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "PRESTAMOS DEL  " + fecha_dos + "  " + NUMERO_DE_CAJA + "')", conexionmy);
                    SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "PRESTAMOS DEL  " + fecha_dos + " ')", conexionmy);
                    leyendaP.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + LOCALIDAD + "  " + "CAJA N " + NUMERO_DE_CAJA + "')", conexionmy);
                    //SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + LOCALIDAD + "  " + "CAJA # " + NUMERO_DE_CAJA + " " + server.Substring(9,5) + "')", conexionmy);
                    leyendaPL.ExecuteNonQuery();
                    conexionmy.Close();

                    poliza_prestamos();
                }


                checkBoxX1.Checked = false;

                crear_poliza();




            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }

        //CREAR EL TABLE DEL POOL DE POLIZA DE INTERES
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
            año_inicial = 2011;//2006
            nom = 1;
            valor = Convert.ToInt32(año);
            cuatro = "0" + mes;
            while (año_inicial < valor)
            {
                nom = nom + 1;
                año_inicial = año_inicial + 1;

            }

            int numero = valor;
            numero = int.Parse(numero.ToString().Substring(3));
            nom = numero;

            if (nom < 10)
            {
                seis = Convert.ToString("00" + nom);
            }
            if (nom >= 10)
            {
                seis = Convert.ToString("0" + nom);
            }

            catorce = CUENTA_PRESTAMOS + "-" + seis + "-" + cuatro;
            DataTable tablapres = new DataTable();
            tablapres.Clear();
            SqlDataAdapter datos_pres = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  contratos INNER JOIN " + caja + " " +
                " ON contratos.contrato = " + caja + ".contrato  " +
                " WHERE contratos.Fechacons= '" + Convert.ToDateTime(fecha_uno).ToString("dd-MM-yyyy") + "' and " + caja + ".concepto LIKE '%PRESTAMO%' " +
                " order by contratos.contrato asc " +
                "", conexionmy);
            datos_pres.Fill(tablapres);

            if (tablapres.Rows.Count == 0)
            {

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
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "PRESTAMO CONTRATO #  " + CONTRATO + "')", conexionmy);
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
                    SqlCommand comando_F = new SqlCommand("USE " + server + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + "CAJA" + " " + NUMERO_DE_CAJA + " CONTRATO #  " + CONTRATO + "')", conexionmy);
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

            }

        }

        //CREAR ARCHIVO POR ARCHIVO POOL DE PRESTAMOS
        private void notas()
        {
            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            DataTable tablaprestamos = new DataTable();
            tablaprestamos.Clear();
            SqlDataAdapter datos_presa = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  prestamos_poliza " +
                " order by  no asc " +
                "", conexionmy);
            datos_presa.Fill(tablaprestamos);

            if (!Directory.Exists(path + "/PRESTAMOS-" + caja))
            {
                Directory.CreateDirectory(path + "/PRESTAMOS-" + caja);
            }

            System.IO.StreamWriter escribir = new System.IO.StreamWriter(path + "/PRESTAMOS-" + caja + "/PRESTAMOS " + caja + "-" + fecha_uno + ".pol");
            foreach (DataRow fila in tablaprestamos.Rows)
            {
                escribir.WriteLine(fila[1]);

            }
            escribir.Close();


        }

        #endregion

        #region Cobros rutina
        //INICIO DE GENERACION DE ARCHIVO INTERES
        public void cobros()
        {
            try
            {

                letra2 = textBoxX2.Text;
                SqlConnection conexionmy = new SqlConnection(sqlcnx);

                //progressBarX2.Maximum = dias_en_mes;
                //progressBarX2.Value = 0;
                //string CONTRATO, STATUS, PRESTAMO;
                int _nveces;
                _nveces = 1;
                if (!string.IsNullOrEmpty(txtDias.Text) && !string.IsNullOrEmpty(txtDiasFinal.Text))
                {

                    if (int.Parse(txtDias.Text) > 0 && int.Parse(txtDiasFinal.Text) > 0)
                    {
                        _nveces = int.Parse(txtDias.Text);
                        dias_en_mes = int.Parse(txtDiasFinal.Text);
                    }
                    else
                    {
                        _nveces = 1;
                    }

                }



                for (nveces2 = _nveces; nveces2 <= dias_en_mes; nveces2++)
                {
                    backgroundWorker3.ReportProgress(nveces2);

                    fecha_cinco = Convert.ToString(nveces2);
                    if (fecha_cinco.Length == 1)
                    {
                        fecha_cinco = "0" + fecha_cinco;
                    }
                    fecha_uno = año + "-" + mes + "-" + fecha_cinco;
                    fecha_dos = fecha_cinco + "/" + _mesLetra + "/" + año;

                    conexionmy.Open();
                    SqlCommand comando_crea = new SqlCommand("USE " + server +
                    " truncate table cobros_poliza " +
                    " ", conexionmy);
                    comando_crea.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand simbolo = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + letra2 + "')", conexionmy);
                    simbolo.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand diadepoliza = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + fecha_cinco + "')", conexionmy);
                    diadepoliza.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    //SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + "COBROS DEL  " + fecha_dos + "  " + NUMERO_DE_CAJA + "')", conexionmy);
                    SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + "COBROS DEL  " + fecha_dos + " ')", conexionmy);
                    leyendaP.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + LOCALIDAD + "  " + "CAJA N " + NUMERO_DE_CAJA + "')", conexionmy);
                    //SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + LOCALIDAD + "  " + "CAJA # " + NUMERO_DE_CAJA + " "+ server.Substring(9,5) +  "')", conexionmy);
                    leyendaPL.ExecuteNonQuery();
                    conexionmy.Close();

                    poliza_cobro();
                }

                checkBoxX2.Checked = false;
                crear_poliza();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        //CREAR EL TABLE DEL POOL DE POLIZA DE INTERES
        public void poliza_cobro()
        {


            SqlConnection conexionmy = new SqlConnection(sqlcnx);

            fecha_tres = año;
            fecha_cuatro = mes;

            fecha_dos = fecha_cinco + "/" + mes + "/" + año;
            int año_inicial, nom, valor;
            string seis = "";
            string cuatro;
            string catorce;
            año_inicial = 2011;//2006
            nom = 1;
            valor = Convert.ToInt32(año);
            cuatro = "0" + mes;
            while (año_inicial < valor)
            {
                nom = nom + 1;
                año_inicial = año_inicial + 1;

            }
            int numero = valor;
            numero = int.Parse(numero.ToString().Substring(3));
            nom = numero;
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
           " SELECT facturas.factura, facturas.contrato, " + caja + ".debe, facturas.importefact, facturas.ivafact, facturas.totalfact, facturas.status , facturas.[total gastos_op] , facturas.GASTOS_OPERACION, facturas.antes_refrendo FROM  facturas " +
                " INNER JOIN " + caja + " ON facturas.factura = " + caja + ".folio WHERE " +
                " facturas.FechaFact='" + Convert.ToDateTime(fecha_uno).ToString("yyyy-dd-MM") + "' and " + caja + ".concepto like '%" + letra + "%' order by facturas.factura asc" +
           "", conexionmy);

            //comienza el if de lleno o vacio
            comando_carga.Fill(tablacobros);

            if (tablacobros.Rows.Count == 0)
            {

            }
            else
            {
                string FACTURA, CONTRATO, DEBE, INTERES, IVA, TOTALFACT, STATUS, FECHACONS, PRESTAMO = "0", quince = "0", TOTAL_GASTOS_OP = "0", GASTOS_OP = "0", ANTES_REFRENDO = "0";
                decimal SumaGastosOperacion = 0;
                foreach (DataRow dr in tablacobros.Rows)
                {


                    FACTURA = dr[0].ToString();//contrato
                    CONTRATO = dr[1].ToString(); //status
                    DEBE = dr[2].ToString(); //prestamo $$
                    INTERES = dr[3].ToString();//contrato
                    IVA = dr[4].ToString(); //status
                    TOTALFACT = dr[5].ToString(); //prestamo $$
                    STATUS = dr[6].ToString();//contratodr.Cells[6].Value.ToString();//contrato

                    TOTAL_GASTOS_OP = dr[7].ToString();// 
                    GASTOS_OP = dr[8].ToString();// 
                    ANTES_REFRENDO = dr[9].ToString();

                    if (!string.IsNullOrEmpty(TOTAL_GASTOS_OP) || !string.IsNullOrWhiteSpace(TOTAL_GASTOS_OP))
                    {
                        if (!string.IsNullOrEmpty(GASTOS_OP) || !string.IsNullOrWhiteSpace(GASTOS_OP))
                        {
                            if (decimal.Parse(TOTALFACT) != (decimal.Parse(TOTAL_GASTOS_OP) + decimal.Parse(ANTES_REFRENDO)))
                            {

                                SumaGastosOperacion += decimal.Parse(TOTAL_GASTOS_OP);

                            }

                         
                        }



                    }


                    //NUMERO DE LA CUENTA
                    conexionmy.Open();
                    SqlCommand comando_A = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + CUENTA_CAJA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    comando_A.ExecuteNonQuery();
                    conexionmy.Close();
                    //LAS LEYENDAS
                    conexionmy.Open();
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + FACTURA + "   " + "C-" + CONTRATO + "')", conexionmy);
                    comando_B.ExecuteNonQuery();
                    conexionmy.Close();
                    //EL MONTO DEL PRESTAMO
                    string monto_valor;

                    decimal unosidebed = decimal.Round(Convert.ToDecimal(DEBE), 2, MidpointRounding.AwayFromZero);
                    monto_valor = string.Format("{0:0.000000}", unosidebed);


                    if (STATUS == "CANCELADO")
                    {
                        conexionmy.Open();
                        SqlCommand comando_C = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        comando_C.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_D = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + monto_valor + ",1.00" + "')", conexionmy);
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


                        año_inicial_uno = 2011;//2006
                        nom_uno = 1;
                        valor_uno = Convert.ToInt32(año_fechacons);//año actual

                        cuatro_uno = "0" + mes_fechacons;

                        while (año_inicial_uno < valor_uno)
                        {
                            nom_uno = nom_uno + 1;
                            año_inicial_uno = año_inicial_uno + 1;

                        }

                        int numero_ = valor_uno;
                        numero_ = int.Parse(numero_.ToString().Substring(3));
                        nom = numero_;

                        if (nom < 10)
                        {
                            seis_uno = Convert.ToString("00" + nom);
                        }
                        if (nom >= 10)
                        {
                            seis_uno = Convert.ToString("0" + nom);
                        }

                        //if (nom_uno < 10)
                        //{
                        //    seis_uno = Convert.ToString("00" + nom_uno);
                        //}
                        //if (nom_uno >= 10)
                        //{
                        //    seis_uno = Convert.ToString("0" + nom_uno);
                        //}
                        quince = CUENTA_PRESTAMOS + "-" + seis_uno + "-" + cuatro_uno;

                    }

                    //

                    //AHORA LA CUENTA DE LA CAJA
                    conexionmy.Open();
                    SqlCommand comando_E = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + quince + "                ," + "  " + DEPARTAMENTO + " ')", conexionmy);
                    comando_E.ExecuteNonQuery();
                    conexionmy.Close();
                    //ahora leyenda caja
                    conexionmy.Open();
                    SqlCommand comando_F = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + FACTURA + "   " + "C-" + CONTRATO + "')", conexionmy);
                    comando_F.ExecuteNonQuery();
                    conexionmy.Close();
                    //inserto espacio
                    //ahora la el monto de prestamo
                    conexionmy.Open();
                    SqlCommand comando_G = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('')", conexionmy);
                    comando_G.ExecuteNonQuery();
                    conexionmy.Close();
                    //nuevamene la cantidad
                    //ahora la el monto de prestamo             

                    string valor_convertido;


                    decimal unosidebede = decimal.Round(Convert.ToDecimal(PRESTAMO), 2, MidpointRounding.AwayFromZero);
                    valor_convertido = string.Format("{0:0.000000}", unosidebede);

                    if (STATUS == "CANCELADO")
                    {
                        conexionmy.Open();
                        SqlCommand comando_H = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        comando_H.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_I = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + valor_convertido + ",1.00" + "')", conexionmy);
                        comando_I.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    //'ahora cuenta caja
                    conexionmy.Open();
                    SqlCommand comando_Ia = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + CUENTA_INTERES + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    comando_Ia.ExecuteNonQuery();
                    conexionmy.Close();

                    //ahora leyenda interes
                    conexionmy.Open();
                    SqlCommand comando_Iaa = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + "INTERESES- " + FACTURA + " " + "C-" + CONTRATO + "')", conexionmy);
                    comando_Iaa.ExecuteNonQuery();
                    conexionmy.Close();


                    //un nuevo espacio
                    conexionmy.Open();
                    SqlCommand comando_Iaabk = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta) VALUES('')", conexionmy);
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
                        SqlCommand comando_Ha = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        comando_Ha.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_Iz = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + valor_convertido3 + ",1.00" + "')", conexionmy);
                        comando_Iz.ExecuteNonQuery();
                        conexionmy.Close();
                    }


                    //CUENTA SI ES QUE EXISTEN GASTOS DE OPERACION NO FACTURADOS
                    //ingreso el registro en la poliza la cuenta de cobros de gastos de operacion

                    if (SumaGastosOperacion > 0)
                    {

                        //string.Format("{0:0.000000}", SumaGastosOperacion);

                        #region TRASH
                        //int año_inicial_uno, nom_uno, valor_uno;
                        //string seis_uno = "";
                        //string cuatro_uno;
                        //string CuentaGastos = "";

                        //string año_fechacons, mes_fechacons;
                        //año_fechacons = Convert.ToDateTime(fecha_uno).ToString("yyyy");
                        //mes_fechacons = Convert.ToDateTime(fecha_uno).ToString("MM");


                        //año_inicial_uno = 2006;
                        //nom_uno = 1;
                        //valor_uno = Convert.ToInt32(año_fechacons);

                        //cuatro_uno = "0" + mes_fechacons;

                        //while (año_inicial_uno < valor_uno)
                        //{
                        //    nom_uno = nom_uno + 1;
                        //    año_inicial_uno = año_inicial_uno + 1;

                        //}
                        //if (nom_uno < 10)
                        //{
                        //    seis_uno = Convert.ToString("00" + nom_uno);
                        //}
                        //if (nom_uno >= 10)
                        //{
                        //    seis_uno = Convert.ToString("0" + nom_uno);
                        //}

                        //CuentaGastos = "1003" + "-" + seis_uno + "-" + cuatro_uno; 
                        #endregion



                        //CUENTA DEL GASTO DE OPERACION
                        conexionmy.Open();
                        SqlCommand comando_GastosOperacion_1 = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + CUENTA_INTERES + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                        comando_GastosOperacion_1.ExecuteNonQuery();
                        conexionmy.Close();



                        //leyenda iva
                        conexionmy.Open();
                        SqlCommand comando_GastosOperacion_2 = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + "GAST OP-" + FACTURA + " " + "C-" + CONTRATO + "')", conexionmy);
                        comando_GastosOperacion_2.ExecuteNonQuery();
                        conexionmy.Close();



                        //ESPACIO
                        conexionmy.Open();
                        SqlCommand comando_GastosOperacion_3 = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta)VALUES('')", conexionmy);
                        comando_GastosOperacion_3.ExecuteNonQuery();
                        conexionmy.Close();


                        conexionmy.Open();
                        SqlCommand comando_GastosOperacion_4 = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + string.Format("{0:0.000000}", SumaGastosOperacion) + ",1.00" + "')", conexionmy);
                        comando_GastosOperacion_4.ExecuteNonQuery();
                        conexionmy.Close();

                        SumaGastosOperacion = 0;


                    }










                    //FIN DE CUENTA SI ES QUE EXISTEN GASTOS DE OPERACION





                    //'ahora la cuenta del iva
                    conexionmy.Open();
                    SqlCommand comando_Iza = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + CUENTA_IVA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    comando_Iza.ExecuteNonQuery();
                    conexionmy.Close();
                    //leyenda iva
                    conexionmy.Open();
                    SqlCommand comando_Iaabe = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + "IVA- " + FACTURA + "   " + "C-" + CONTRATO + "')", conexionmy);
                    comando_Iaabe.ExecuteNonQuery();
                    conexionmy.Close();
                    //un nuevo espacio
                    conexionmy.Open();
                    SqlCommand comando_Iaaba = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta)VALUES('')", conexionmy);
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
                        SqlCommand comando_Ha = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        comando_Ha.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_Iz = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + valor_convertido35 + ",1.00" + "')", conexionmy);
                        comando_Iz.ExecuteNonQuery();
                        conexionmy.Close();
                    }

                    //
                }//cierre del ciclo for each




                conexionmy.Open();
                SqlCommand comando_J = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('FIN')", conexionmy);
                comando_J.ExecuteNonQuery();
                conexionmy.Close();
                notas_cobro();
            }


        }
        //CREAR ARCHIVO POR ARCHVIVO POOL DE INTERES
        public void notas_cobro()
        {
            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            //conexionmy.ConnectionString = sqlcnx;//conexion mysql
            DataTable tablaprestamos = new DataTable();
            tablaprestamos.Clear();
            SqlDataAdapter datos_presa = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  cobros_poliza " +
                " order by  no asc " +
                "", conexionmy);
            datos_presa.Fill(tablaprestamos);



            if (!Directory.Exists(path + "/NOTADEPAGO-" + caja))
            {
                Directory.CreateDirectory(path + "/NOTADEPAGO-" + caja);
            }
            System.IO.StreamWriter escribir = new System.IO.StreamWriter(path + "/NOTADEPAGO-" + caja + "/NOTADEPAGO " + caja + "-" + fecha_uno + ".pol");
            foreach (DataRow fila in tablaprestamos.Rows)
            {
                escribir.WriteLine(fila[1]);
            }
            escribir.Close();

        }
        #endregion

        #region Interes Diario //ya no operan
        //INICIO DE GENERACION DE ARCHIVO INTERES DIARIO
        public void diario_interes()
        {
            try
            {
                letra3 = textBoxX4.Text;

                SqlConnection conexionmy = new SqlConnection(sqlcnx);


                progressBarX3.Maximum = dias_en_mes;
                progressBarX3.Value = 0;
                //string CONTRATO, STATUS, PRESTAMO;
                int _nveces;
                _nveces = 1;
                if (!string.IsNullOrEmpty(txtDias.Text) && !string.IsNullOrEmpty(txtDiasFinal.Text))
                {

                    if (int.Parse(txtDias.Text) > 0 && int.Parse(txtDiasFinal.Text) > 0)
                    {
                        _nveces = int.Parse(txtDias.Text);
                        dias_en_mes = int.Parse(txtDiasFinal.Text);
                    }
                    else
                    {
                        _nveces = 1;
                    }

                }

                for (nveces3 = _nveces; nveces3 <= dias_en_mes; nveces3++)
                {

                    backgroundWorker4.ReportProgress(nveces3);

                    fecha_cinco = Convert.ToString(nveces3);
                    if (fecha_cinco.Length == 1)
                    {
                        fecha_cinco = "0" + fecha_cinco;
                    }
                    fecha_uno = año + "-" + mes + "-" + fecha_cinco;
                    fecha_dos = fecha_cinco + "/" + _mesLetra + "/" + año;

                    conexionmy.Open();
                    SqlCommand comando_crea = new SqlCommand("USE " + server +
                    " truncate table diario_poliza " +
                    " ", conexionmy);
                    comando_crea.ExecuteNonQuery();
                    conexionmy.Close();
                    //MessageBox.Show("" + fecha_cinco);
                    conexionmy.Open();
                    SqlCommand simbolo = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + letra3 + "')", conexionmy);//letra
                    simbolo.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand diadepoliza = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + fecha_cinco + "')", conexionmy);//dia
                    diadepoliza.ExecuteNonQuery();
                    conexionmy.Close();
                    ////este es diferente con iff segn la empresa
                    //LAS LEYENDAS con if segun empr
                    if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                    {
                        conexionmy.Open();
                        //SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Interes Diario " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + "')", conexionmy);
                        SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Interes Diario " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + " " + server.Substring(9, 5) + "')", conexionmy);
                        leyendaP.ExecuteNonQuery();
                        conexionmy.Close();

                    }
                    else if (empresa_Conta == "MONTE ROS SA DE CV")
                    {
                        conexionmy.Open();
                        //SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Interes Diario " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + "')", conexionmy);
                        SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Interes Diario " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + " " + server.Substring(9, 5) + "')", conexionmy);
                        leyendaP.ExecuteNonQuery();
                        conexionmy.Close();

                    }
                    else
                    {
                        conexionmy.Open();
                        //SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Interes Diario " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + "')", conexionmy);
                        SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Interes Diario " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + " " + server.Substring(9, 5) + "')", conexionmy);
                        leyendaP.ExecuteNonQuery();
                        conexionmy.Close();

                    }

                    conexionmy.Open();
                    SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('.')", conexionmy);
                    leyendaPL.ExecuteNonQuery();
                    conexionmy.Close();

                    poliza_diario();
                }

                checkBoxX3.Checked = false;
                crear_poliza();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //CREAR EL TABLE DEL POOL DE POLIZA DE INTERES DIARIO
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
            SqlDataAdapter comando_carga = new SqlDataAdapter("USE " + server +
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
                iva_solo = tablacobros.Rows[0][1].ToString();// tablacobros.Rows[0].ItemArray[1].ToString();//status

                if (importe_del_iva == "")
                {
                    importe_del_iva = "0.00";
                    iva_solo = "0.00";


                }


                //NUMERO DE LA CUENTA
                conexionmy.Open();
                SqlCommand comando_A = new SqlCommand("USE " + server + " INSERT into diario_poliza(cuenta) VALUES('" + CUENTA_INTERES + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                comando_A.ExecuteNonQuery();
                conexionmy.Close();


                //LAS LEYENDAS con if segun empr
                if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + "')", conexionmy);
                    comando_B.ExecuteNonQuery();
                    conexionmy.Close();
                }
                else if (empresa_Conta == "MONTE ROS SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_B.ExecuteNonQuery();
                    conexionmy.Close();
                }
                else
                {
                    conexionmy.Open();
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + "')", conexionmy);
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
                SqlCommand comando_D = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + monto_valor_T + ",1.00" + "')", conexionmy);
                comando_D.ExecuteNonQuery();
                conexionmy.Close();

                conexionmy.Open();

                //ACTUALIZACION DE CUENTA DE INTERES DIARIO
                string CadenaOriginal = CUENTA_INTERES.Trim();
                string primerElemento = CUENTA_INTERES.Substring(0, 5);
                string segundoElemento = "001";
                string tercerElemento = CUENTA_INTERES.Substring(8, 4);
                string cadenaConvertida = primerElemento + segundoElemento + tercerElemento;

                //02 FEBRERO 2019
                SqlCommand comando_D_D = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + cadenaConvertida + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                comando_D_D.ExecuteNonQuery();
                conexionmy.Close();

                //LAS LEYENDAS2 segun la empresa entr otr if
                if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_B_b = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + "')", conexionmy);
                    comando_B_b.ExecuteNonQuery();
                    conexionmy.Close();
                }
                else if (empresa_Conta == "MONTE ROS SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_B_b = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_B_b.ExecuteNonQuery();
                    conexionmy.Close();
                }
                else
                {
                    conexionmy.Open();
                    SqlCommand comando_B_b = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_B_b.ExecuteNonQuery();
                    conexionmy.Close();
                }


                conexionmy.Open();
                SqlCommand comando_B_e = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES(' ')", conexionmy);
                comando_B_e.ExecuteNonQuery();
                conexionmy.Close();

                string monto_valor_G;
                // monto_valor_G = Convert.ToString(string.Format("{0:F2}", importe_del_iva)); //Convert.ToString(importe_del_iva);
                //monto_valor_G = string.Format(monto_valor_G, "####0.000000");

                decimal unosidebe = decimal.Round(Convert.ToDecimal(importe_del_iva), 2, MidpointRounding.AwayFromZero);
                monto_valor_G = string.Format("{0:0.000000}", unosidebe);

                conexionmy.Open();
                SqlCommand comando_D_H = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + monto_valor_G + ",1.00" + "')", conexionmy);
                comando_D_H.ExecuteNonQuery();
                conexionmy.Close();

                conexionmy.Open();
                SqlCommand comando_D_R = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + CUENTA_IVA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                comando_D_R.ExecuteNonQuery();
                conexionmy.Close();
                //if segun la empresa
                //LAS LEYENDAS2 segun la empresa entr otr if
                if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_D_S = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + "')", conexionmy);
                    comando_D_S.ExecuteNonQuery();
                    conexionmy.Close();

                }
                else if (empresa_Conta == "MONTE ROS SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_D_S = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_D_S.ExecuteNonQuery();
                    conexionmy.Close();

                }
                else
                {
                    conexionmy.Open();
                    SqlCommand comando_D_S = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + "')", conexionmy);
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
                SqlCommand comando_D_S_R = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + monto_valor_GG + ",1.00" + "')", conexionmy);
                comando_D_S_R.ExecuteNonQuery();
                conexionmy.Close();

                conexionmy.Open();
                SqlCommand comando_D_S_Ru = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + CUENTA_IVA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                comando_D_S_Ru.ExecuteNonQuery();
                conexionmy.Close();


                //LAS LEYENDAS2 segun la empresa entr otr if
                if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_D_Sw = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "- " + NUMERO_DE_CAJA + " " + mes_letra + " " + fecha_cinco + "')", conexionmy);
                    comando_D_Sw.ExecuteNonQuery();
                    conexionmy.Close();

                }
                else if (empresa_Conta == "MONTE ROS SA DE CV")
                {
                    conexionmy.Open();
                    SqlCommand comando_D_Sw = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_D_Sw.ExecuteNonQuery();
                    conexionmy.Close();

                }
                else
                {
                    conexionmy.Open();
                    SqlCommand comando_D_Sw = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + "Notas de Pago " + LUGAR_CONTA + "-" + NUMERO_DE_CAJA + "')", conexionmy);
                    comando_D_Sw.ExecuteNonQuery();
                    conexionmy.Close();

                }
                conexionmy.Open();
                SqlCommand comando_D_S_F = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES(' ')", conexionmy);
                comando_D_S_F.ExecuteNonQuery();
                conexionmy.Close();

                string monto_valor_GGe;
                //monto_valor_GGe = Convert.ToString(string.Format("{0:F2}", iva_solo));// Convert.ToString(iva_solo);
                //monto_valor_GGe = string.Format(monto_valor_GGe, "####0.000000");
                decimal unosidebede = decimal.Round(Convert.ToDecimal(iva_solo), 2, MidpointRounding.AwayFromZero);
                monto_valor_GGe = string.Format("{0:0.000000}", unosidebede);

                conexionmy.Open();
                SqlCommand comando_D_S_Ra = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('" + monto_valor_GGe + ",1.00" + "')", conexionmy);
                comando_D_S_Ra.ExecuteNonQuery();
                conexionmy.Close();

                conexionmy.Open();
                SqlCommand comando_D_S_Fi = new SqlCommand("USE " + server + " INSERT INTO diario_poliza(cuenta)VALUES('FIN')", conexionmy);
                comando_D_S_Fi.ExecuteNonQuery();
                conexionmy.Close();
                notas_diario();
            }

        }
        //CREAR ARCHIVO POR ARCHIVO POOL DE INTERES DIARIO
        public void notas_diario()
        {
            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            //conexionmy.ConnectionString = sqlcnx;//conexion mysql
            DataTable tablaprestamos = new DataTable();
            tablaprestamos.Clear();
            SqlDataAdapter datos_presa = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  diario_poliza " +
                " order by  no asc " +
                "", conexionmy);
            datos_presa.Fill(tablaprestamos);
            //this.dataGridView2.DataSource = tablaprestamos;


            if (!Directory.Exists(path + "/DIARIO-" + caja))
            {
                Directory.CreateDirectory(path + "/DIARIO-" + caja);
            }


            System.IO.StreamWriter escribir = new System.IO.StreamWriter(path + "/DIARIO-" + caja + "/DIARIO " + caja + "-" + fecha_uno + ".pol");
            foreach (DataRow fila in tablaprestamos.Rows)
            {
                escribir.WriteLine(fila[1]);
            }
            escribir.Close();

        }



        #endregion


        #region Interes Semanal

        //CREAR EL TABLE DEL POOL DE POLIZA DE INTERES SEMANAL
        public void poliza_interes_semanal()
        {
            try
            {
                // int dia_en_mes_ret;
                int mes_configurado;
                int total_dias;
                string primer_fecha;
                string segunda_fecha;
                int contador = 1;
                string que_dia_es;
                int suma = 0;
                int resta = 0;

                DateTime fecha_lunes;
                DateTime fecha_domingo;
                DateTime fecha_dia_final;
                DateTime fecha_dia_inicial;
                //mes
                //año


                total_dias = DateTime.DaysInMonth(Convert.ToInt32(año), Convert.ToInt32(mes));
                mes_configurado = Convert.ToInt32(mes);
                fecha_dia_final = Convert.ToDateTime(año + "-" + mes + "-" + total_dias);
                fecha_dia_inicial = Convert.ToDateTime(año + "-" + mes + "-" + "01");
                string mi_Caja;
                string total_fact = "0";
                string total_iva_fact = "0";
                //2. cargo las cajas de la sucursal seleccionada
                //progressBarX4.Maximum = dias_en_mes;
                //progressBarX4.Value = 0;


                if (!string.IsNullOrEmpty(txtDias.Text) && !string.IsNullOrEmpty(txtDiasFinal.Text))
                {
                    if (int.Parse(txtDias.Text) > 0 && int.Parse(txtDiasFinal.Text) > 0)
                    {
                        total_dias = int.Parse(txtDiasFinal.Text);


                        if (txtDiasFinal.Text.Length == 1)
                        {
                            txtDiasFinal.Text = "0" + txtDiasFinal.Text;
                        }
                        if (txtDias.Text.Length == 1)
                        {
                            txtDias.Text = "0" + txtDias.Text;
                        }
                        fecha_dia_final = Convert.ToDateTime(año + "-" + mes + "-" + txtDiasFinal.Text);
                        fecha_dia_inicial = Convert.ToDateTime(año + "-" + mes + "-" + txtDias.Text);
                        contador = total_dias;
                    }

                }


                while (contador <= total_dias)//segun el numero de dias del mes
                {
                    backgroundWorker5.ReportProgress(contador);

                    fecha_cinco = Convert.ToString(contador);
                    if (fecha_cinco.Length == 1)
                    {
                        fecha_cinco = "0" + fecha_cinco;
                    }

                    SqlConnection conexionmy = new SqlConnection(sqlcnx);
                    //conexionmy.ConnectionString = sqlcnx;//conexion mysql

                    conexionmy.Open();
                    SqlCommand comando_crea = new SqlCommand("USE " + server +
                    " truncate table interesSemanal_poliza " +
                    " ", conexionmy);
                    comando_crea.ExecuteNonQuery();
                    conexionmy.Close();
                    //MessageBox.Show("" + fecha_cinco);
                    conexionmy.Open();
                    SqlCommand simbolo = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + letra4 + "')", conexionmy);//letra
                    simbolo.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand diadepoliza = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + fecha_cinco + "')", conexionmy);//dia
                    diadepoliza.ExecuteNonQuery();
                    conexionmy.Close();

                  
                    ////este es diferente con iff segn la empresa

                    //obtengo la primer fecha


                    primer_fecha = año + "-" + mes + "-" + contador;
                    segunda_fecha = año + "-" + mes + "-" + contador;



                    //obtengo el dia lunes
                    que_dia_es = Convert.ToDateTime(primer_fecha).DayOfWeek.ToString("d");
                    //verifico que dias
                    fecha_lunes = DateTime.Parse(primer_fecha);


                    if (Convert.ToInt32(que_dia_es) == 1)//es lunes?
                    {
                        //inicio = 1;//no voy a restar

                        //ahora mis inicios
                        fecha_lunes = DateTime.Parse(primer_fecha);
                    }
                    else
                    {
                        //ahora mis inicios
                        resta = Convert.ToInt32(que_dia_es) - 1;//dias que restare para llegar al lunes
                        fecha_lunes = DateTime.Parse(primer_fecha).AddDays(-resta);
                    }




                    ////y aqui empieza la primer compracion saber si aun estoy en el mes que corresponde de dia lunes
                    int MES_INICIO = fecha_lunes.Month;

                    if (MES_INICIO == mes_configurado && Convert.ToInt32(Convert.ToDateTime(fecha_lunes).DayOfWeek.ToString("d")) == 1)//si mes inicio es igual a mes operativo
                    {
                        //no hago nada y se que si esta el lunes en ese dia de la semna
                    }
                    else
                    {
                        fecha_lunes = fecha_lunes.AddDays(1);
                    }


                    if (fecha_lunes.Month < mes_configurado)
                    {
                        fecha_lunes = fecha_dia_inicial;
                    }

                    //ERROR DE ENERO 1 SEMANA
                    if (fecha_lunes.Year < int.Parse(año))
                    {
                        fecha_lunes = fecha_dia_inicial;
                    }



                    //domingo
                    if (Convert.ToInt32(que_dia_es) == 0)//si es domingo
                    {
                        fecha_domingo = DateTime.Parse(segunda_fecha);
                    }
                    else
                    {
                        suma = 7 - Convert.ToInt32(que_dia_es);//seran dias que sumare para llegar a domindngo
                        fecha_domingo = DateTime.Parse(segunda_fecha).AddDays(suma);
                    }

                    int MES_FINAL = fecha_domingo.Month;
                    if (MES_FINAL == mes_configurado && Convert.ToInt32(Convert.ToDateTime(fecha_domingo).DayOfWeek.ToString("d")) == 0 && fecha_domingo > fecha_lunes)//si mes inicio es igual a mes operativo
                    {
                        //no hago nada y se que si esta el lunes en ese dia de la semna
                    }
                    else
                    {
                        fecha_domingo = fecha_domingo.AddDays(suma);
                    }
                    if (fecha_domingo == fecha_lunes) //se entiende que domingo es igual a 7 y entonces sumo
                    {
                        fecha_domingo = fecha_domingo.AddDays(6);

                    }

                    if (fecha_domingo.Month != mes_configurado)
                    {
                        fecha_domingo = fecha_dia_final;
                    }

                    if (fecha_lunes > fecha_domingo)
                    {
                        fecha_lunes = fecha_domingo;
                    }

                    switch (mes)
                    {
                        case "01":
                            mes_letra = "Ene.";
                            break;
                        case "02":
                            mes_letra = "Feb.";
                            break;
                        case "03":
                            mes_letra = "Mar.";
                            break;
                        case "04":
                            mes_letra = "Abr";
                            break;
                        case "05":
                            mes_letra = "May.";
                            break;
                        case "06":
                            mes_letra = "Jun.";
                            break;
                        case "07":
                            mes_letra = "Jul.";
                            break;
                        case "08":
                            mes_letra = "Ago.";
                            break;
                        case "09":
                            mes_letra = "Sep.";
                            break;
                        case "10":
                            mes_letra = "Oct.";
                            break;
                        case "11":
                            mes_letra = "Nov.";
                            break;
                        case "12":
                            mes_letra = "Dic.";
                            break;
                        default:
                            break;

                    }


                    conexionmy.Open();

                    SqlCommand diadepolizas = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Interes Semanal" + " " + server.Substring(9, 5) + " "
                                                                    + Convert.ToDateTime(fecha_lunes).ToString("dd/MMM/yyyy") + " al " + Convert.ToDateTime(fecha_domingo).ToString("dd/MMM/yyyy") + "')", conexionmy);//dia
                    //SqlCommand diadepolizas = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Desempeños " + LUGAR_CONTA + " " + NUMERO_DE_CAJA + " " + server.Substring(9, 5) + "')", conexionmy);//dia
                    diadepolizas.ExecuteNonQuery();
                    conexionmy.Close();


                    conexionmy.Open();
                    SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('')", conexionmy);
                    leyendaPL.ExecuteNonQuery();
                    conexionmy.Close();








                    foreach (DataRow dr in tablecaja_dos.Rows)
                    {

                        mi_Caja = dr[0].ToString();//contrato

                        DataTable tablasql = new DataTable();
                        SqlDataAdapter datosSql = new SqlDataAdapter("USE SEMP2013_NAU_1 " +
                          "Select *  from contabilidad where leyenda2='" + mi_Caja + "'" +
                          "", conexionmy);
                        datosSql.Fill(tablasql);
                        //SqlDataAdapter datosSql = new SqlDataAdapter("USE " + server + "  " +
                        //    "Select *  from contabilidad where leyenda2='" + mi_Caja + "'" +
                        //    "", conexionmy);
                        //datosSql.Fill(tablasql);
                        LOCALIDAD = tablasql.Rows[0].ItemArray[2].ToString();
                        CUENTA_CAJA = tablasql.Rows[0].ItemArray[6].ToString();
                        CUENTA_INTERES = tablasql.Rows[0].ItemArray[7].ToString();
                        CUENTA_IVA = tablasql.Rows[0].ItemArray[8].ToString();
                        NUMERO_DE_CAJA = tablasql.Rows[0].ItemArray[3].ToString();
                        DEPARTAMENTO = tablasql.Rows[0].ItemArray[9].ToString();
                        CUENTA_PRESTAMOS = tablasql.Rows[0].ItemArray[5].ToString();
                        TABLA_CAJA = tablasql.Rows[0].ItemArray[11].ToString();

                        //string letra = "DESEMP";
                        DataTable tablacobros_Y = new DataTable();
                        tablacobros_Y.Clear();
                        //  MessageBox.Show("" + fecha_lunes.ToString("yyyy-MM-dd 00:00:00") +
                        // fecha_domingo.ToString("yyyy-MM-dd 00:00:00") );
                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);

                        if (!string.IsNullOrEmpty(txtDias.Text) && !string.IsNullOrEmpty(txtDiasFinal.Text))
                        {
                            if (int.Parse(txtDias.Text) > 0 && int.Parse(txtDiasFinal.Text) > 0)
                            {
                                if (txtDiasFinal.Text.Length == 1)
                                {
                                    txtDiasFinal.Text = "0" + txtDiasFinal.Text;
                                }
                                if (txtDias.Text.Length == 1)
                                {
                                    txtDias.Text = "0" + txtDias.Text;
                                }
                                fecha_lunes = Convert.ToDateTime(año + "-" + mes + "-" + txtDias.Text);
                                fecha_domingo = Convert.ToDateTime(año + "-" + mes + "-" + txtDiasFinal.Text);

                            }

                        }


                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);



                        SqlDataAdapter comando_carga_Y = new SqlDataAdapter("USE " + server +
                        " SELECT sum(ImporteFact), sum(IVAFACT) FROM  facturas INNER JOIN " + mi_Caja + " " +
                        " ON facturas.factura = " + mi_Caja + ".folio AND facturas.STATUS = 'VALIDO'  WHERE " +
                        " facturas.FechaFact BETWEEN  '" + fecha_lunes.ToString("dd-MM-yyyy") + "' and " +
                        " '" + fecha_domingo.ToString("dd-MM-yyyy") + "' AND " + mi_Caja + ".concepto like '%DESEMP%'  " +
                       "", conexionmy);
                        tablacobros_Y.Clear();

                        //comienza el if de lleno o vacio
                        comando_carga_Y.Fill(tablacobros_Y);
                        // this.dataGridView1.DataSource = tablacobros_Y;
                        //veo si esta vacio el dia
                        if (tablacobros_Y.Rows.Count == 0)
                        {

                            //exportar con nota de no hubo operacion este dia y seguir con rutina
                        }
                        else
                        {

                            string importe_del_iva, iva_solo;



                            importe_del_iva = tablacobros_Y.Rows[0][0].ToString();//this.dataGridView1.Rows[0].ItemArray[0].ToString();//contrato
                            iva_solo = tablacobros_Y.Rows[0][1].ToString(); //tablacobros_Y.Rows[0].ItemArray[1].ToString();//status

                            if (importe_del_iva == "")
                            {
                                importe_del_iva = "0.00";
                                iva_solo = "0.00";


                            }
                            // MessageBox.Show("" + importe_del_iva);
                            // MessageBox.Show("" + iva_solo);

                            //NUMERO DE LA CUENTA
                            conexionmy.Open();
                            SqlCommand comando_A = new SqlCommand("USE " + server + " INSERT into interesSemanal_poliza(cuenta) VALUES('" + CUENTA_INTERES + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_A.ExecuteNonQuery();
                            conexionmy.Close();


                            //LAS LEYENDAS con if segun empr
                            conexionmy.Open();
                            SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Interes " + server.Substring(9, 5)  + " Caja" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + "')", conexionmy);
                           // SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Desempeños " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                            comando_B.ExecuteNonQuery();
                            conexionmy.Close();



                            //EL MONTO DEL PRESTAMO me quede donde dice monteo del importe en notas d diarii linea 3236 en vb net
                            string monto_valor_T;
                            //monto_valor_T = Convert.ToString(string.Format("{0:F2}", importe_del_iva)); //Convert.ToString(importe_del_iva);
                            //monto_valor_T = string.Format(monto_valor_T, "####0.000000");

                            decimal unos = decimal.Round(Convert.ToDecimal(importe_del_iva), 2, MidpointRounding.AwayFromZero);
                            monto_valor_T = string.Format("{0:0.000000}", unos);



                            conexionmy.Open();
                            SqlCommand comando_D = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + monto_valor_T + ",1.00" + "')", conexionmy);
                            comando_D.ExecuteNonQuery();
                            conexionmy.Close();

                            conexionmy.Open();
                            SqlCommand comando_D_D = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + CUENTA_IVA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_D_D.ExecuteNonQuery();
                            conexionmy.Close();

                            //LAS LEYENDAS2 segun la empresa entr otr if
                            conexionmy.Open();
                            //SqlCommand comando_Bw = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Desempeños " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                            SqlCommand comando_Bw = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "IVA " + server.Substring(9, 5) + " Caja" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + "')", conexionmy);
                            comando_Bw.ExecuteNonQuery();
                            conexionmy.Close();



                            string monto_valor_G;
                            //monto_valor_G = Convert.ToString(string.Format("{0:F2}", iva_solo)); //Convert.ToString(iva_solo);
                            //monto_valor_G = string.Format(monto_valor_G, "####0.000000");

                            decimal unosi = decimal.Round(Convert.ToDecimal(iva_solo), 2, MidpointRounding.AwayFromZero);
                            monto_valor_G = string.Format("{0:0.000000}", unosi);


                            conexionmy.Open();
                            SqlCommand comando_D_H = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + monto_valor_G + ",1.00" + "')", conexionmy);
                            comando_D_H.ExecuteNonQuery();
                            conexionmy.Close();



                            //string monto_valor_GG;
                            //monto_valor_GG = Convert.ToString(iva_solo);
                            //monto_valor_GG = string.Format(monto_valor_GG, "####0.000000");

                            //conexionmy.Open();
                            //SqlCommand comando_D_S_R = new SqlCommand("USE " + this.label4.Text + " INSERT INTO prestamos_poliza(cuenta)VALUES('" + monto_valor_GG + ",1.00" + "')", conexionmy);
                            //comando_D_S_R.ExecuteNonQuery();
                            //conexionmy.Close();

                        }//cierre del else
                    }//cierre del for each
                     //string letra = "DESEMP";
                    DataTable tablacobros_L = new DataTable();
                    tablacobros_L.Clear();
                    SqlDataAdapter comando_carga_L = new SqlDataAdapter("USE " + server +
                    " SELECT sum(ImporteFact), sum(IVAFACT) FROM  facturas   WHERE STATUS = 'VALIDO' and " +
                    " FechaFact BETWEEN  '" + fecha_lunes.ToString("dd-MM-yyyy") + "' and  '" + fecha_domingo.ToString("dd-MM-yyyy") + "'" +
                   "", conexionmy);
                    //comienza el if de lleno o vacio
                    comando_carga_L.Fill(tablacobros_L);
                    //this.dataGridView4.DataSource = tablacobros_L;
                    //veo si esta vacio el dia
                    if (tablacobros_L.Rows.Count == 0)
                    {
                    }
                    else
                    {
                        total_fact = tablacobros_L.Rows[0].ItemArray[0].ToString();//contrato
                        total_iva_fact = tablacobros_L.Rows[0].ItemArray[1].ToString();//status

                        if (total_fact == "")
                        {
                            total_fact = "0.00";
                            total_iva_fact = "0.00";


                        }
                        


                        if (empresa_Conta == "DDR GARCIA SA DE CV")
                        {
                            conexionmy.Open();
                            SqlCommand comando_D_S_T = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "401-001-001-001" + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_D_S_T.ExecuteNonQuery();
                            conexionmy.Close();

                            conexionmy.Open();
                            //SqlCommand comando_Bwe = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Desempeños " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                            SqlCommand comando_Bwe = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Interes " + server.Substring(9, 5) + "')", conexionmy);
                            comando_Bwe.ExecuteNonQuery();
                            conexionmy.Close();
                        }
                        else if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                        {
                            conexionmy.Open();
                            SqlCommand comando_D_S_T = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "401-001-001-001" + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_D_S_T.ExecuteNonQuery();
                            conexionmy.Close();

                            conexionmy.Open();
                            //SqlCommand comando_Bwe = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Desempeños " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                            SqlCommand comando_Bwe = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Interes " + server.Substring(9, 5) + "')", conexionmy);
                            comando_Bwe.ExecuteNonQuery();
                            conexionmy.Close();

                        }
                        else if (empresa_Conta == "MONTE ROS SA DE CV")
                        {
                            conexionmy.Open();
                            SqlCommand comando_D_S_T = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "401-001-001-001" + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_D_S_T.ExecuteNonQuery();
                            conexionmy.Close();

                            conexionmy.Open();
                            //SqlCommand comando_Bwe = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Desempeños " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                            SqlCommand comando_Bwe = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Interes " + server.Substring(9, 5) + "')", conexionmy);
                            comando_Bwe.ExecuteNonQuery();
                            conexionmy.Close();

                        }


                        ////LAS LEYENDAS2 segun la empresa entr otr if
                        //if (empresa_Conta == "DDR GARCIA SA DE CV")
                        //{

                        //}
                        //else if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                        //{


                        //}



                        conexionmy.Open();
                        SqlCommand comando_B_re = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('')", conexionmy);
                        comando_B_re.ExecuteNonQuery();
                        conexionmy.Close();

                        string monto_valor_Ger;
                        //monto_valor_Ger = Convert.ToString(string.Format("{0:F2}", total_fact)); //Convert.ToString(total_fact);
                        //monto_valor_Ger = string.Format(monto_valor_Ger, "####0.000000");

                        decimal unosid = decimal.Round(Convert.ToDecimal(total_fact), 2, MidpointRounding.AwayFromZero);
                        monto_valor_Ger = string.Format("{0:0.000000}", unosid);


                        conexionmy.Open();
                        SqlCommand comando_D_Hi = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + monto_valor_Ger + ",1.00" + "')", conexionmy);
                        comando_D_Hi.ExecuteNonQuery();
                        conexionmy.Close();

                     

                        if (empresa_Conta == "DDR GARCIA SA DE CV")
                        {
                            conexionmy.Open();
                            SqlCommand comando_D_Hiw = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "213-001-001-001" + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_D_Hiw.ExecuteNonQuery();
                            conexionmy.Close();

                            conexionmy.Open();
                            //SqlCommand comando_Bwek = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Desempeños " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                            SqlCommand comando_Bwek = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Iva " + server.Substring(9, 5) + "')", conexionmy);
                            comando_Bwek.ExecuteNonQuery();
                            conexionmy.Close();
                        }
                        else if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                        {
                            conexionmy.Open();
                            SqlCommand comando_D_Hiw = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "213-001-001-001" + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_D_Hiw.ExecuteNonQuery();
                            conexionmy.Close();

                            conexionmy.Open();
                            //SqlCommand comando_Bwek = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Desempeños " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                            SqlCommand comando_Bwek = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Iva " + server.Substring(9, 5) + "')", conexionmy);
                            comando_Bwek.ExecuteNonQuery();
                            conexionmy.Close();

                        }
                        else if (empresa_Conta == "MONTE ROS SA DE CV")
                        {
                            conexionmy.Open();
                            SqlCommand comando_D_Hiw = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "213-001-001-001" + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_D_Hiw.ExecuteNonQuery();
                            conexionmy.Close();

                            conexionmy.Open();
                            //SqlCommand comando_Bwek = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Desempeños " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                            SqlCommand comando_Bwek = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + "Iva " + server.Substring(9, 5) + "')", conexionmy);
                            comando_Bwek.ExecuteNonQuery();
                            conexionmy.Close();

                        }
                        conexionmy.Open();
                        SqlCommand comando_B_rem = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('')", conexionmy);
                        comando_B_rem.ExecuteNonQuery();
                        conexionmy.Close();

                        string monto_valor_Gero;
                        //monto_valor_Gero = Convert.ToString(string.Format("{0:F2}", total_iva_fact)); //Convert.ToString(total_iva_fact);
                        //monto_valor_Gero = string.Format(monto_valor_Gero, "####0.000000");

                        decimal unoside = decimal.Round(Convert.ToDecimal(total_iva_fact), 2, MidpointRounding.AwayFromZero);
                        monto_valor_Gero = string.Format("{0:0.000000}", unoside);

                        conexionmy.Open();
                        SqlCommand comando_D_Hiu = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('" + monto_valor_Gero + ",1.00" + "')", conexionmy);
                        comando_D_Hiu.ExecuteNonQuery();
                        conexionmy.Close();

                        conexionmy.Open();
                        SqlCommand comando_B_remu = new SqlCommand("USE " + server + " INSERT INTO interesSemanal_poliza(cuenta)VALUES('FIN')", conexionmy);
                        comando_B_remu.ExecuteNonQuery();
                        conexionmy.Close();
                    }//cierre del else total




                    //LA CONSULTA FIN

                    fecha_inicial = fecha_lunes.ToString("dd-MM-yyyy");
                    fecha_final = fecha_domingo.ToString("dd-MM-yyyy");





                    notas_interes_semanal();

                    contador = contador + 1;



                }//while

                checkBoxX4.Checked = false;
                crear_poliza();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //CREAR ARCHIVO POR ARCHIVO POOL DE INTERES SEMANAL
        public void notas_interes_semanal()
        {
            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            // conexionmy.ConnectionString = sqlcnx;//conexion mysql
            DataTable tablaprestamos = new DataTable();
            tablaprestamos.Clear();
            SqlDataAdapter datos_presa = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  interesSemanal_poliza " +
                " order by  no asc " +
                "", conexionmy);
            datos_presa.Fill(tablaprestamos);
            // this.dataGridView2.DataSource = tablaprestamos;
            if (!Directory.Exists(path + "/intSEM-" + caja))
            {
                Directory.CreateDirectory(path + "/intSEM-" + caja);
            }

            if (File.Exists(path + "/intSEM-" + caja + "/intSEM " + caja + "-" + fecha_inicial + "-" + fecha_final + ".pol")) { }
            else
            {
                System.IO.StreamWriter escribir = new System.IO.StreamWriter(path + "/intSEM-" + caja + "/intSEM " + caja + "-" + fecha_inicial + "-" + fecha_final + ".pol");
                foreach (DataRow fila in tablaprestamos.Rows)
                {
                    escribir.WriteLine(fila[1]);
                }
                escribir.Close();
            }

        }

        #endregion

        #region Remision Semanal

        //CREAR EL TABLE DEL POOL DE POLIZA DE REMISION SEMANAL
        public void poliza_remisiones_semanal()
        {
            try
            {
                // int dia_en_mes_ret;
                int mes_configurado;
                int total_dias;
                string primer_fecha;
                string segunda_fecha;
                int contador = 1;
                string que_dia_es;
                int suma = 0;
                int resta = 0;

                DateTime fecha_lunes;
                DateTime fecha_domingo;
                DateTime fecha_dia_final;
                DateTime fecha_dia_inicial;
                //mes
                //año
                total_dias = DateTime.DaysInMonth(Convert.ToInt32(año), Convert.ToInt32(mes));
                mes_configurado = Convert.ToInt32(mes);
                fecha_dia_final = Convert.ToDateTime(año + "-" + mes + "-" + total_dias);
                fecha_dia_inicial = Convert.ToDateTime(año + "-" + mes + "-" + "01");
                string mi_Caja;
                string total_rem = "0";
                //  string total_iva_fact = "0";
                //2. cargo las cajas de la sucursal seleccionada
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

                //progressBarX5.Maximum = dias_en_mes;
                //progressBarX5.Value = 0;
                if (!string.IsNullOrEmpty(txtDias.Text) && !string.IsNullOrEmpty(txtDiasFinal.Text))
                {
                    if (int.Parse(txtDias.Text) > 0 && int.Parse(txtDiasFinal.Text) > 0)
                    {
                        total_dias = int.Parse(txtDiasFinal.Text);


                        if (txtDiasFinal.Text.Length == 1)
                        {
                            txtDiasFinal.Text = "0" + txtDiasFinal.Text;
                        }
                        if (txtDias.Text.Length == 1)
                        {
                            txtDias.Text = "0" + txtDias.Text;
                        }
                        fecha_dia_final = Convert.ToDateTime(año + "-" + mes + "-" + txtDiasFinal.Text);
                        fecha_dia_inicial = Convert.ToDateTime(año + "-" + mes + "-" + txtDias.Text);
                        contador = total_dias;
                    }
                }


                //////////////////////////////////////////////

                while (contador <= total_dias)//segun el numero de dias del mes
                {
                    backgroundWorker6.ReportProgress(contador);

                    fecha_cinco = Convert.ToString(contador);
                    if (fecha_cinco.Length == 1)
                    {
                        fecha_cinco = "0" + fecha_cinco;
                    }

                    SqlConnection conexionmy = new SqlConnection();
                    conexionmy.ConnectionString = sqlcnx;//conexion mysql

                    conexionmy.Open();
                    SqlCommand comando_crea = new SqlCommand("USE " + server +
                    " truncate table remisionSemanal_poliza " +
                    " ", conexionmy);
                    comando_crea.ExecuteNonQuery();
                    conexionmy.Close();
                    //MessageBox.Show("" + fecha_cinco);
                    conexionmy.Open();
                    SqlCommand simbolo = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + letra5 + "')", conexionmy);//letra
                    simbolo.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand diadepoliza = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + fecha_cinco + "')", conexionmy);//dia
                    diadepoliza.ExecuteNonQuery();
                    conexionmy.Close();

                  
                    //obtengo la primer fecha
                    primer_fecha = año + "-" + mes + "-" + contador;
                    segunda_fecha = año + "-" + mes + "-" + contador;
                    //obtengo el dia lunes
                    que_dia_es = Convert.ToDateTime(primer_fecha).DayOfWeek.ToString("d");
                    //verifico que dias

                    if (Convert.ToInt32(que_dia_es) == 1)//es lunes?
                    {
                        //inicio = 1;//no voy a restar

                        //ahora mis inicios
                        fecha_lunes = DateTime.Parse(primer_fecha);
                    }
                    else
                    {
                        //ahora mis inicios
                        resta = Convert.ToInt32(que_dia_es) - 1;//dias que restare para llegar al lunes
                        fecha_lunes = DateTime.Parse(primer_fecha).AddDays(-resta);
                    }

                    ////y aqui empieza la primer compracion saber si aun estoy en el mes que corresponde de dia lunes
                    int MES_INICIO = fecha_lunes.Month;
                    // MessageBox.Show("me inicio" + MES_INICIO + "" + mes_configurado);
                    if (MES_INICIO == mes_configurado && Convert.ToInt32(Convert.ToDateTime(fecha_lunes).DayOfWeek.ToString("d")) == 1)//si mes inicio es igual a mes operativo
                    {
                        //no hago nada y se que si esta el lunes en ese dia de la semna
                    }
                    else
                    {
                        fecha_lunes = fecha_lunes.AddDays(1);
                    }
                    if (fecha_lunes.Month < mes_configurado)
                    {
                        fecha_lunes = fecha_dia_inicial;
                    }

                    //domingo
                    if (Convert.ToInt32(que_dia_es) == 0)//si es domingo
                    {
                        fecha_domingo = DateTime.Parse(segunda_fecha);
                    }
                    else
                    {
                        suma = 7 - Convert.ToInt32(que_dia_es);//seran dias que sumare para llegar a domindngo
                        fecha_domingo = DateTime.Parse(segunda_fecha).AddDays(suma);
                    }
                    int MES_FINAL = fecha_domingo.Month;
                    if (MES_FINAL == mes_configurado && Convert.ToInt32(Convert.ToDateTime(fecha_domingo).DayOfWeek.ToString("d")) == 0 && fecha_domingo > fecha_lunes)//si mes inicio es igual a mes operativo
                    {
                        //no hago nada y se que si esta el lunes en ese dia de la semna
                    }
                    else
                    {
                        fecha_domingo = fecha_domingo.AddDays(suma);
                    }
                    if (fecha_domingo == fecha_lunes) //se entiende que domingo es igual a 7 y entonces sumo
                    {
                        fecha_domingo = fecha_domingo.AddDays(6);

                    }



                    //if (fecha_lunes.Month < mes_configurado)
                    //{
                    //    fecha_lunes = fecha_dia_inicial;
                    //}

                    //ERROR DE ENERO 1 SEMANA
                    if (fecha_lunes.Year < int.Parse(año))
                    {
                        fecha_lunes = fecha_dia_inicial;
                    }






                    if (fecha_domingo.Month != mes_configurado)
                    {
                        fecha_domingo = fecha_dia_final;
                    }

                    if (fecha_lunes > fecha_domingo)
                    {
                        fecha_lunes = fecha_domingo;
                    }



                    switch (mes)
                    {
                        case "01":
                            mes_letra = "Ene.";
                            break;
                        case "02":
                            mes_letra = "Feb.";
                            break;
                        case "03":
                            mes_letra = "Mar.";
                            break;
                        case "04":
                            mes_letra = "Abr.";
                            break;
                        case "05":
                            mes_letra = "May.";
                            break;
                        case "06":
                            mes_letra = "Jun.";
                            break;
                        case "07":
                            mes_letra = "Jul.";
                            break;
                        case "08":
                            mes_letra = "Ago.";
                            break;
                        case "09":
                            mes_letra = "Sep.";
                            break;
                        case "10":
                            mes_letra = "Oct.";
                            break;
                        case "11":
                            mes_letra = "Nov.";
                            break;
                        case "12":
                            mes_letra = "Dic.";
                            break;
                        default:
                            break;

                    }


                    conexionmy.Open();
                    //SqlCommand diadepolizas = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "Remision Semanal" + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + " " + server.Substring(9, 5) + "')", conexionmy);//dia
                    SqlCommand diadepolizas = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "Ventas Semanal" + " " + server.Substring(9, 5) + " "
                                                                    + Convert.ToDateTime(fecha_lunes).ToString("dd/MMM/yyyy") + " al " + Convert.ToDateTime(fecha_domingo).ToString("dd/MMM/yyyy") + "')", conexionmy);//dia
                    diadepolizas.ExecuteNonQuery();
                    conexionmy.Close();


                    conexionmy.Open();
                    SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES(' ')", conexionmy);
                    leyendaPL.ExecuteNonQuery();
                    conexionmy.Close();
                    ////este es diferente con iff segn la empresa



                    foreach (DataRow dr in tablecaja_dos.Rows)
                    {
                        string leyenda_1;
                        mi_Caja = dr[0].ToString();//contrato


                        DataTable tablasql = new DataTable();
                        SqlDataAdapter datosSql = new SqlDataAdapter("USE SEMP2013_NAU_1 " +
                          "Select *  from contabilidad where leyenda2='" + mi_Caja + "'" +
                          "", conexionmy);
                        datosSql.Fill(tablasql);
                        //SqlDataAdapter datosSql = new SqlDataAdapter("USE " + server + "  " +
                        //    "Select *  from contabilidad where leyenda2='" + mi_Caja + "'" +
                        //    "", conexionmy);
                        //datosSql.Fill(tablasql);
                        LOCALIDAD = tablasql.Rows[0].ItemArray[2].ToString();
                        CUENTA_CAJA = tablasql.Rows[0].ItemArray[6].ToString();
                        CUENTA_INTERES = tablasql.Rows[0].ItemArray[7].ToString();
                        CUENTA_IVA = tablasql.Rows[0].ItemArray[8].ToString();
                        NUMERO_DE_CAJA = tablasql.Rows[0].ItemArray[3].ToString();
                        DEPARTAMENTO = tablasql.Rows[0].ItemArray[9].ToString();
                        CUENTA_PRESTAMOS = tablasql.Rows[0].ItemArray[5].ToString();
                        TABLA_CAJA = tablasql.Rows[0].ItemArray[11].ToString();
                        leyenda_1 = tablasql.Rows[0].ItemArray[10].ToString();
                        //string letra = "DESEMP";
                        DataTable tablacobros_Y = new DataTable();
                        tablacobros_Y.Clear();


                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);
                        if (!string.IsNullOrEmpty(txtDias.Text) && !string.IsNullOrEmpty(txtDiasFinal.Text))
                        {
                            if (int.Parse(txtDias.Text) > 0 && int.Parse(txtDiasFinal.Text) > 0)
                            {
                                if (txtDiasFinal.Text.Length == 1)
                                {
                                    txtDiasFinal.Text = "0" + txtDiasFinal.Text;
                                }
                                if (txtDias.Text.Length == 1)
                                {
                                    txtDias.Text = "0" + txtDias.Text;
                                }
                                fecha_lunes = Convert.ToDateTime(año + "-" + mes + "-" + txtDias.Text);
                                fecha_domingo = Convert.ToDateTime(año + "-" + mes + "-" + txtDiasFinal.Text);

                            }



                        }

                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);
                        //MessageBox.Show("" + mi_Caja);



                        SqlDataAdapter comando_carga_Y = new SqlDataAdapter("USE " + server +
                            " SELECT SUM(Importe) AS Expr1" +
                   " FROM remisiones " +
                  " WHERE     (Fecha BETWEEN  '" + fecha_lunes.ToString("dd-MM-yyyy") + "' AND '" + fecha_domingo.ToString("dd-MM-yyyy") + "') " +
                  " AND (caja = '" + leyenda_1 + "') AND (status = 'VENDIDO') OR" +
                  " (Fecha BETWEEN  '" + fecha_lunes.ToString("dd-MM-yyyy") + "' AND '" + fecha_domingo.ToString("dd-MM-yyyy") + "') " +
                  " AND (caja ='" + leyenda_1 + "') AND (status = 'PAGADO')" +
                    "", conexionmy);
                        tablacobros_Y.Clear();

                        //comienza el if de lleno o vacio
                        comando_carga_Y.Fill(tablacobros_Y);
                        //  this.dataGridView1.DataSource = tablacobros_Y;
                        //veo si esta vacio el dia
                        if (tablacobros_Y.Rows.Count == 0)
                        {

                            //exportar con nota de no hubo operacion este dia y seguir con rutina
                        }
                        else
                        {

                            string importe_remision;

                            importe_remision = tablacobros_Y.Rows[0][0].ToString();//this.dataGridView1.Rows[0].ItemArray[0].ToString();//contrato
                            if (importe_remision == "")
                            {
                                importe_remision = "0.00";

                            }
                            //NUMERO DE LA CUENTA
                            conexionmy.Open();//'" & Me.TextBox11.Text + "                " + ",  " + Me.TextBox15.Text & "'
                            SqlCommand comando_A = new SqlCommand("USE " + server + " INSERT into remisionSemanal_poliza(cuenta) VALUES('" + CUENTA_CAJA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_A.ExecuteNonQuery();
                            conexionmy.Close();


                            //LAS LEYENDAS con if segun empr
                            conexionmy.Open();
                            //SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "Ventas " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                            SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "Ventas " + server.Substring(9, 5) + " Caja" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + "')", conexionmy);

                            comando_B.ExecuteNonQuery();
                            conexionmy.Close();



                            //EL MONTO DEL PRESTAMO me quede donde dice monteo del importe en notas d diarii linea 3236 en vb net
                            string monto_valor_T;
                            //  importe_remision = Convert.ToString(string.Format("{0:F2}",importe_remision));
                            // monto_valor_T = Convert.ToString(string.Format("{0:000000}", importe_remision)); //Convert.ToString(string.Format("####0.000000",monto_valor_T));

                            decimal uno = decimal.Round(Convert.ToDecimal(importe_remision), 2, MidpointRounding.AwayFromZero);
                            monto_valor_T = string.Format("{0:0.000000}", uno);


                            conexionmy.Open();
                            SqlCommand comando_D = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + monto_valor_T + ",1.00" + "')", conexionmy);
                            comando_D.ExecuteNonQuery();
                            conexionmy.Close();


                        }//cierre del else
                    }//cierre del for each
                    //string letra = "DESEMP";
                    DataTable tablacobros_L = new DataTable();
                    tablacobros_L.Clear();
                    SqlDataAdapter comando_carga_L = new SqlDataAdapter(" USE " + server +
                            " SELECT SUM(Importe) AS Expr1" +
                   " FROM remisiones " +
                  " WHERE     (Fecha BETWEEN  '" + fecha_lunes.ToString("dd-MM-yyyy") + "' AND '" + fecha_domingo.ToString("dd-MM-yyyy") + "') " +
                  " AND (status = 'VENDIDO') OR" +
                  " (Fecha BETWEEN  '" + fecha_lunes.ToString("dd-MM-yyyy") + "' AND '" + fecha_domingo.ToString("dd-MM-yyyy") + "') " +
                  " AND  (status = 'PAGADO')" +
                    "", conexionmy);
                    //comienza el if de lleno o vacio
                    comando_carga_L.Fill(tablacobros_L);
                    //this.dataGridView4.DataSource = tablacobros_L;
                    //veo si esta vacio el dia
                    if (tablacobros_L.Rows.Count == 0)
                    {
                    }
                    else
                    {
                        total_rem = tablacobros_L.Rows[0].ItemArray[0].ToString();//contrato
                        if (total_rem == "")
                        {
                            total_rem = "0.00";

                        }

                        int año_inicial = 2011;//2006
                        int nom = 1;
                        int valor = Convert.ToInt32(año);
                        string cuatro = "0" + mes;
                        string seis = "";
                        while (año_inicial < valor)
                        {
                            nom = nom + 1;
                            año_inicial = año_inicial + 1;

                        }
                        int numero = valor;
                        numero = int.Parse(valor.ToString().Substring(3));
                        nom = numero;
                        //TODO:Fecha
                        if (nom < 10)
                        {
                            seis = Convert.ToString("00" + nom);
                        }
                        if (nom >= 10)
                        {
                            seis = Convert.ToString("0" + nom);
                        }
                        //TODO:If cuenta 1003
                        string leyenda_cuenta = string.Empty;
                       // leyenda_cuenta = "106-101" + "-" + seis + "-" + cuatro;
                        if (empresa_Conta == "DDR GARCIA SA DE CV")
                        {
                            leyenda_cuenta = "106-101" + "-" + seis + "-" + cuatro;
                        }
                        else if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                        {
                            if (LOCALIDAD == "TLA_2")
                            {
                                leyenda_cuenta = "106-101" + "-" + seis + "-" + cuatro;

                            }
                            if (LOCALIDAD == "TLX_2")
                            {
                                leyenda_cuenta = "106-102" + "-" + seis + "-" + cuatro;

                            }
                            if (LOCALIDAD == "MIX_1")
                            {
                                leyenda_cuenta = "106-103" + "-" + seis + "-" + cuatro;

                            }

                        }
                        else if (empresa_Conta == "MONTE ROS SA DE CV")
                        {
                            if (LOCALIDAD == "TLA_3")
                            {
                                leyenda_cuenta = "106-101" + "-" + seis + "-" + cuatro;

                            }
                            if (LOCALIDAD == "PRG_1")
                            {
                                leyenda_cuenta = "106-102" + "-" + seis + "-" + cuatro;

                            }
                           

                        }



                        conexionmy.Open();
                        SqlCommand comando_D_S_T = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + leyenda_cuenta + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                        comando_D_S_T.ExecuteNonQuery();
                        conexionmy.Close();

                        conexionmy.Open();
                        //SqlCommand comando_Bwe = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "Remisiones " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                        SqlCommand comando_Bwe = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "Ventas " + server.Substring(9, 5) + " Caja" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + "')", conexionmy);
                        comando_Bwe.ExecuteNonQuery();
                        conexionmy.Close();


                        conexionmy.Open();
                        SqlCommand comando_B_re = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES(' ')", conexionmy);
                        comando_B_re.ExecuteNonQuery();
                        conexionmy.Close();

                        //ahora la el monto de importe del 97.5%


                        double importe_rem = Convert.ToDouble(total_rem);
                        decimal total_valor;
                        double porcentaje = 100.00 - double.Parse(txtPorcentaje.Text);//  97.5;//antes 97.5//ajuste de tomar el 90% 19 dic 2018
                                                                                      // decimal porcentajeConvertido=10
                        total_valor = Convert.ToDecimal(importe_rem * porcentaje / 100);


                        string monto_valor_Te;
                        // total_valor = Convert.ToString(string.Format("{0:F2}", total_valor)); //Convert.ToString(total_valor);
                        //  monto_valor_Te = Convert.ToString(string.Format("{0:000000}",total_valor)); //string.Format(monto_valor_Te, "####0.000000");

                        total_valor = decimal.Round(total_valor, 2, MidpointRounding.AwayFromZero);
                        monto_valor_Te = string.Format("{0:0.000000}", total_valor);


                        conexionmy.Open();
                        SqlCommand comando_B_rea = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + monto_valor_Te + ",1.00" + "')", conexionmy);
                        comando_B_rea.ExecuteNonQuery();
                        conexionmy.Close();




                        if (empresa_Conta == "DDR GARCIA SA DE CV")
                        {

                            conexionmy.Open();
                            SqlCommand comando_D_Hiw = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "401-001-001-002" + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_D_Hiw.ExecuteNonQuery();
                            conexionmy.Close();
                        }
                        else if (empresa_Conta == "COMERCIAL INTERMODAL SA DE CV")
                        {
                            conexionmy.Open();
                            SqlCommand comando_D_Hiw = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "401-001-001-002" + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_D_Hiw.ExecuteNonQuery();
                            conexionmy.Close();

                        }
                        else if (empresa_Conta == "MONTE ROS SA DE CV")
                        {

                            conexionmy.Open();
                            SqlCommand comando_D_Hiw = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "401-001-001-002" + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                            comando_D_Hiw.ExecuteNonQuery();
                            conexionmy.Close();

                        }







                        // ('" + "Remisiones " + server.Substring(9, 5) + " Caja" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + "')", conexionmy);
                        conexionmy.Open();
                        //SqlCommand comando_Bweh = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "Ventas " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                        SqlCommand comando_Bweh = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "Ventas " + server.Substring(9, 5) + " Caja" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + "')", conexionmy);
                        comando_Bweh.ExecuteNonQuery();
                        conexionmy.Close();

                        conexionmy.Open();
                        SqlCommand comando_B_remn = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES(' ')", conexionmy);
                        comando_B_remn.ExecuteNonQuery();
                        conexionmy.Close();


                        double diferencia_total;
                        decimal subtotal_sin_iva;
                        diferencia_total = Convert.ToDouble(total_rem) - Convert.ToDouble(total_valor);
                        subtotal_sin_iva = Convert.ToDecimal(diferencia_total / 1.16);

                        string monto_valor_Gero;
                        // subtotal_sin_iva = Convert.ToString(string.Format("{0:F2}", subtotal_sin_iva));//Convert.ToString(subtotal_sin_iva);
                        // monto_valor_Gero = Convert.ToString(string.Format("{0:F6}", subtotal_sin_iva)); //string.Format(monto_valor_Gero, "####0.000000");

                        subtotal_sin_iva = decimal.Round(subtotal_sin_iva, 2, MidpointRounding.AwayFromZero);
                        monto_valor_Gero = string.Format("{0:0.000000}", subtotal_sin_iva);



                        conexionmy.Open();
                        SqlCommand comando_D_Hiu = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + monto_valor_Gero + ",1.00" + "')", conexionmy);
                        comando_D_Hiu.ExecuteNonQuery();
                        conexionmy.Close();



                        conexionmy.Open();
                        SqlCommand comando_D_Hiwi = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "213-001-001-001" + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                        comando_D_Hiwi.ExecuteNonQuery();
                        conexionmy.Close();




                        // ('" + "Remisiones " + server.Substring(9, 5) + " Caja" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + "')", conexionmy);
                        conexionmy.Open();
                       // SqlCommand comando_Bwehe = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "Ventas " + LUGAR_CONTA + " C" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + " " + fecha_cinco + "')", conexionmy);
                        SqlCommand comando_Bwehe = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + "Ventas " + server.Substring(9, 5) + " Caja" + NUMERO_DE_CAJA + " " + mes_letra + " " + año + "')", conexionmy);


                        comando_Bwehe.ExecuteNonQuery();
                        conexionmy.Close();

                        conexionmy.Open();
                        SqlCommand comando_B_remne = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES(' ')", conexionmy);
                        comando_B_remne.ExecuteNonQuery();
                        conexionmy.Close();

                        string iva_total;
                        iva_total = Convert.ToString(Convert.ToDouble(total_rem) - Convert.ToDouble(total_valor) - Convert.ToDouble(subtotal_sin_iva));

                        decimal iva_T;
                        string monto_valor_Geroa;
                        iva_T = decimal.Round(Convert.ToDecimal(iva_total), 2, MidpointRounding.AwayFromZero);
                        monto_valor_Geroa = string.Format("{0:0.000000}", iva_T);

                        conexionmy.Open();
                        SqlCommand comando_D_Hius = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('" + monto_valor_Geroa + ",1.00" + "')", conexionmy);
                        comando_D_Hius.ExecuteNonQuery();
                        conexionmy.Close();

                        conexionmy.Open();
                        SqlCommand comando_B_remnea = new SqlCommand("USE " + server + " INSERT INTO remisionSemanal_poliza(cuenta)VALUES('FIN')", conexionmy);
                        comando_B_remnea.ExecuteNonQuery();
                        conexionmy.Close();

                    }//cierre del else total




                    //LA CONSULTA FIN

                    fecha_inicial = fecha_lunes.ToString("dd-MM-yyyy");
                    fecha_final = fecha_domingo.ToString("dd-MM-yyyy");





                    notas_remision_semanal();

                    contador = contador + 1;



                }//while

                checkBoxX5.Checked = false;
                crear_poliza();
            }//try
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        //CREAR EL TABLE DEL POOL DE POLIZA DE REMISION SEMANAL
        public void notas_remision_semanal()
        {
            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            //conexionmy.ConnectionString = sqlcnx;//conexion mysql
            DataTable tablaprestamos = new DataTable();
            tablaprestamos.Clear();
            SqlDataAdapter datos_presa = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  remisionSemanal_poliza " +
                " order by  no asc " +
                "", conexionmy);
            datos_presa.Fill(tablaprestamos);
            //this.dataGridView2.DataSource = tablaprestamos;
            if (!Directory.Exists(path + "/REMSEM-" + caja))
            {
                Directory.CreateDirectory(path + "/REMSEM-" + caja);
            }

            if (File.Exists(path + "/REMSEM-" + caja + "/REMSEM " + caja + "-" + fecha_inicial + "-" + fecha_final + ".pol")) { }
            else
            {
                System.IO.StreamWriter escribir = new System.IO.StreamWriter(path + "/REMSEM-" + caja + "/REMSEM " + caja + "-" + fecha_inicial + "-" + fecha_final + ".pol");
                foreach (DataRow fila in tablaprestamos.Rows)
                {
                    escribir.WriteLine(fila[1]);
                }
                escribir.Close();
            }


        }
        #endregion


        #region Contratos Desempeñados
        //INICIO DE GENERACION DE ARCHIVO INTERES
        public void ContratosDesempeñados()
        {
            try
            {
                //caja = this.comboBoxEx4.Text;
                letra6 = textBoxX7.Text;
                SqlConnection conexionmy = new SqlConnection(sqlcnx);


                //obtengo datos de la caja

                DataTable tablasql = new DataTable();
                //SqlDataAdapter datosSql = new SqlDataAdapter("USE " + server + "  " +
                //    "Select *  from contabilidad where leyenda2='" + caja + "'" +
                //    "", conexionmy);
                SqlDataAdapter datosSql = new SqlDataAdapter("USE SEMP2013_NAU_1  " +
                    "Select *  from contabilidad where leyenda2='" + caja + "'" +
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
                //progressBarX2.Maximum = dias_en_mes;
                //progressBarX2.Value = 0;
                //string CONTRATO, STATUS, PRESTAMO;
                int _nveces;
                _nveces = 1;
                if (!string.IsNullOrEmpty(txtDias.Text) && !string.IsNullOrEmpty(txtDiasFinal.Text))
                {

                    if (int.Parse(txtDias.Text) > 0 && int.Parse(txtDiasFinal.Text) > 0)
                    {
                        _nveces = int.Parse(txtDias.Text);
                        dias_en_mes = int.Parse(txtDiasFinal.Text);
                    }
                    else
                    {
                        _nveces = 1;
                    }



                }


                for (nveces2 = _nveces; nveces2 <= dias_en_mes; nveces2++)
                {
                    backgroundWorker7.ReportProgress(nveces2);

                    fecha_cinco = Convert.ToString(nveces2);
                    if (fecha_cinco.Length == 1)
                    {
                        fecha_cinco = "0" + fecha_cinco;
                    }
                    fecha_uno = año + "-" + mes + "-" + fecha_cinco;
                    fecha_dos = fecha_cinco + "/" + mes + "/" + año;

                    conexionmy.Open();
                    SqlCommand comando_crea = new SqlCommand("USE " + server +
                    " truncate table cobros_poliza_desempenos " +
                    " ", conexionmy);
                    comando_crea.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand simbolo = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('" + letra6 + "')", conexionmy);
                    simbolo.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand diadepoliza = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('" + fecha_cinco + "')", conexionmy);
                    diadepoliza.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('" + "CONTRATOS DESEMPEÑADOS DEL  " + fecha_dos + "  " + NUMERO_DE_CAJA + " ')", conexionmy);
                    //SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('" + "CONTRATOS DESEMPENADOS DEL  " + fecha_dos + "  " + NUMERO_DE_CAJA + " " + server.Substring(9,5) + "')", conexionmy);
                    leyendaP.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('" + server.Substring(9, 5) +  "  " + "Caja " + NUMERO_DE_CAJA + "')", conexionmy);
                    leyendaPL.ExecuteNonQuery();
                    conexionmy.Close();

                    poliza_contratos_desempeñados();
                }

                checkBoxX6.Checked = false;
                crear_poliza();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        //CREAR EL TABLE DEL POOL DE POLIZA DE INTERES
        public void poliza_contratos_desempeñados()
        {


            SqlConnection conexionmy = new SqlConnection(sqlcnx);

            //obtengo datos de la caja




            fecha_tres = año;
            fecha_cuatro = mes;

            fecha_dos = fecha_cinco + "/" + mes + "/" + año;
            int año_inicial, nom, valor;
            string seis = "";
            string cuatro;
            string catorce;
            año_inicial = 2011;//2006
            nom = 1;
            valor = Convert.ToInt32(año);
            cuatro = "0" + mes;
            while (año_inicial < valor)
            {
                nom = nom + 1;
                año_inicial = año_inicial + 1;

            }

            int numero_ = valor;
            numero_ = int.Parse(numero_.ToString().Substring(3));
            nom = numero_;

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
                " SELECT b.Status, b.Prestamo, b.Contrato, A.folio, b.FechaCons	 FROM  contratos as b " +
                " INNER JOIN " + caja + " as A ON b.Contrato = a.Contrato WHERE " +
                " a.Fecha='" + Convert.ToDateTime(fecha_uno).ToString("yyyy-dd-MM") + "' and a.concepto like '%" + letra + "%' order by a.mov asc" +
                 "", conexionmy);

            //comienza el if de lleno o vacio
            comando_carga.Fill(tablacobros);

            if (tablacobros.Rows.Count == 0)
            {

            }
            else
            {
                string FACTURA, CONTRATO, DEBE, STATUS, FECHACONS, PRESTAMO = "0", quince = "0";
                foreach (DataRow dr in tablacobros.Rows)
                {


                    FACTURA = dr[3].ToString();//FOLIO (NO REQUERIDO)
                    CONTRATO = dr[2].ToString(); //CONTRATO
                    DEBE = dr[1].ToString(); //PRESTAMO
                    PRESTAMO = DEBE;
                    STATUS = dr[0].ToString();//ESTATUS
                    FECHACONS = dr[4].ToString();//FECHACONS
                    //NUMERO DE LA CUENTA
                    conexionmy.Open();
                    SqlCommand comando_A = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('" + CUENTA_CAJA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    comando_A.ExecuteNonQuery();
                    conexionmy.Close();
                    //LAS LEYENDAS
                    conexionmy.Open();
                    SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('C-" + CONTRATO + "')", conexionmy);
                    comando_B.ExecuteNonQuery();
                    conexionmy.Close();
                    //EL MONTO DEL PRESTAMO
                    string monto_valor;

                    decimal unosidebed = decimal.Round(Convert.ToDecimal(DEBE), 2, MidpointRounding.AwayFromZero);
                    monto_valor = string.Format("{0:0.000000}", unosidebed);


                    if (STATUS == "CANCELADO")
                    {
                        //conexionmy.Open();
                        //SqlCommand comando_C = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                        //comando_C.ExecuteNonQuery();
                        //conexionmy.Close();
                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_D = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('" + monto_valor + ",1.00" + "')", conexionmy);
                        comando_D.ExecuteNonQuery();
                        conexionmy.Close();
                    }

                    //ahora busco le fecha de origen del contrato en cuestion

                    int año_inicial_uno, nom_uno, valor_uno;
                    string seis_uno = "";
                    string cuatro_uno;

                    string año_fechacons, mes_fechacons;
                    año_fechacons = Convert.ToDateTime(FECHACONS).ToString("yyyy");
                    mes_fechacons = Convert.ToDateTime(FECHACONS).ToString("MM");


                    año_inicial_uno = 2011;//2006
                    nom_uno = 1;
                    valor_uno = Convert.ToInt32(año_fechacons);//año actual

                    cuatro_uno = "0" + mes_fechacons;

                    while (año_inicial_uno < valor_uno)
                    {
                        nom_uno = nom_uno + 1;
                        año_inicial_uno = año_inicial_uno + 1;

                    }

                    int numero = valor_uno;
                    numero = int.Parse(numero.ToString().Substring(3));
                    nom_uno = numero;

                    if (nom_uno < 10)
                    {
                        seis_uno = Convert.ToString("00" + nom_uno);
                    }
                    if (nom_uno >= 10)
                    {
                        seis_uno = Convert.ToString("0" + nom_uno);
                    }

                    quince = CUENTA_PRESTAMOS + "-" + seis_uno + "-" + cuatro_uno;

                    //

                    //AHORA LA CUENTA DE LA CAJA
                    conexionmy.Open();
                    SqlCommand comando_E = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('" + quince + "                ," + "  " + DEPARTAMENTO + " ')", conexionmy);
                    comando_E.ExecuteNonQuery();
                    conexionmy.Close();
                    //ahora leyenda caja
                    conexionmy.Open();
                    SqlCommand comando_F = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('C-" + CONTRATO + "')", conexionmy);
                    comando_F.ExecuteNonQuery();
                    conexionmy.Close();
                    //inserto espacio
                    //ahora la el monto de prestamo
                    conexionmy.Open();
                    SqlCommand comando_G = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('')", conexionmy);
                    comando_G.ExecuteNonQuery();
                    conexionmy.Close();
                    //nuevamene la cantidad
                    //ahora la el monto de prestamo             

                    string valor_convertido;


                    decimal unosidebede = decimal.Round(Convert.ToDecimal(DEBE), 2, MidpointRounding.AwayFromZero);
                    valor_convertido = string.Format("{0:0.000000}", unosidebede);

                    if (STATUS == "CANCELADO")
                    {

                    }
                    else
                    {
                        conexionmy.Open();
                        SqlCommand comando_I = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('" + valor_convertido + ",1.00" + "')", conexionmy);
                        comando_I.ExecuteNonQuery();
                        conexionmy.Close();
                    }
                    ////'ahora cuenta caja
                    //conexionmy.Open();
                    //SqlCommand comando_Ia = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + CUENTA_INTERES + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    //comando_Ia.ExecuteNonQuery();
                    //conexionmy.Close();

                    ////ahora leyenda interes
                    //conexionmy.Open();
                    //SqlCommand comando_Iaa = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('C-" + CONTRATO + "')", conexionmy);
                    //comando_Iaa.ExecuteNonQuery();
                    //conexionmy.Close();


                    //un nuevo espacio
                    //conexionmy.Open();
                    //SqlCommand comando_Iaaba = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta)VALUES('')", conexionmy);
                    //comando_Iaaba.ExecuteNonQuery();
                    //conexionmy.Close();


                    //
                }//cierre del ciclo for each
                conexionmy.Open();
                SqlCommand comando_J = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_desempenos(cuenta)VALUES('FIN')", conexionmy);
                comando_J.ExecuteNonQuery();
                conexionmy.Close();
                poliza_contratos_desempeñados_crear();
            }


        }
        //CREAR ARCHIVO POR ARCHVIVO POOL DE INTERES
        public void poliza_contratos_desempeñados_crear()
        {
            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            //conexionmy.ConnectionString = sqlcnx;//conexion mysql
            DataTable tablaprestamos = new DataTable();
            tablaprestamos.Clear();
            SqlDataAdapter datos_presa = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  cobros_poliza_desempenos " +
                " order by  no asc " +
                "", conexionmy);
            datos_presa.Fill(tablaprestamos);



            if (!Directory.Exists(path + "/CONTS_DESEMP-" + caja))
            {
                Directory.CreateDirectory(path + "/CONTS_DESEMP-" + caja);
            }
            System.IO.StreamWriter escribir = new System.IO.StreamWriter(path + "/CONTS_DESEMP-" + caja + "/CONTS_DESEMP " + caja + "-" + fecha_uno + ".pol");
            foreach (DataRow fila in tablaprestamos.Rows)
            {
                escribir.WriteLine(fila[1]);
            }
            escribir.Close();

        }
        #endregion


        #region Solo Interes e Iva x caja
        //INICIO DE GENERACION DE ARCHIVO INTERES
        public void SoloInteresEivaCaja()
        {
            try
            {
                // caja = this.comboBoxEx4.Text;
                letra7 = textBoxX8.Text;
                SqlConnection conexionmy = new SqlConnection(sqlcnx);

                DataTable tablasql = new DataTable();
                SqlDataAdapter datosSql = new SqlDataAdapter("USE SEMP2013_NAU_1  " +
                    "Select *  from contabilidad where leyenda2='" + caja + "'" +
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

                int _nveces;
                _nveces = 1;
                if (!string.IsNullOrEmpty(txtDias.Text) && !string.IsNullOrEmpty(txtDiasFinal.Text))
                {

                    if (int.Parse(txtDias.Text) > 0 && int.Parse(txtDiasFinal.Text) > 0)
                    {
                        _nveces = int.Parse(txtDias.Text);
                        dias_en_mes = int.Parse(txtDiasFinal.Text);
                    }
                    else
                    {
                        _nveces = 1;
                    }

                }


                for (nveces2 = _nveces; nveces2 <= dias_en_mes; nveces2++)
                {
                    backgroundWorker8.ReportProgress(nveces2);

                    fecha_cinco = Convert.ToString(nveces2);
                    if (fecha_cinco.Length == 1)
                    {
                        fecha_cinco = "0" + fecha_cinco;
                    }
                    fecha_uno = año + "-" + mes + "-" + fecha_cinco;
                    fecha_dos = fecha_cinco + "/" + mes + "/" + año;

                    conexionmy.Open();
                    SqlCommand comando_crea = new SqlCommand("USE " + server +
                    " truncate table cobros_poliza_interes " +
                    " ", conexionmy);
                    comando_crea.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand simbolo = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta)VALUES('" + letra7 + "')", conexionmy);
                    simbolo.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand diadepoliza = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta)VALUES('" + fecha_cinco + "')", conexionmy);
                    diadepoliza.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand leyendaP = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta)VALUES('" + "INTERES DEL  " + fecha_dos + "  " + NUMERO_DE_CAJA + " ')", conexionmy);
                    leyendaP.ExecuteNonQuery();
                    conexionmy.Close();

                    conexionmy.Open();
                    SqlCommand leyendaPL = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta)VALUES('" + server.Substring(9, 5) + "  " + "Caja " + NUMERO_DE_CAJA + "')", conexionmy);
                    leyendaPL.ExecuteNonQuery();
                    conexionmy.Close();

                    poliza_SoloInteresEivaCaja();
                }

                checkBoxX7.Checked = false;
                crear_poliza();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        //CREAR EL TABLE DEL POOL DE POLIZA DE INTERES
        public void poliza_SoloInteresEivaCaja()
        {
            

            SqlConnection conexionmy = new SqlConnection(sqlcnx);

            fecha_tres = año;
            fecha_cuatro = mes;

            fecha_dos = fecha_cinco + "/" + mes + "/" + año;
            int año_inicial, nom, valor;
            string seis = "";
            string cuatro;
            string catorce;
            año_inicial = 2011;//2006
            nom = 1;
            valor = Convert.ToInt32(año);
            cuatro = "0" + mes;
            while (año_inicial < valor)
            {
                nom = nom + 1;
                año_inicial = año_inicial + 1;

            }

            int numero = valor;
            numero = int.Parse(valor.ToString().Substring(3));
            nom = numero;

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

            if (tablacobros.Rows.Count == 0)
            {

            }
            else
            {
                string FACTURA, CONTRATO, DEBE, INTERES, IVA, TOTALFACT, STATUS, FECHACONS, PRESTAMO = "0", quince = "0";
                foreach (DataRow dr in tablacobros.Rows)
                {


                    FACTURA = dr[0].ToString();//NOTA DE PAGO
                    CONTRATO = dr[1].ToString(); //CONTRATO
                    DEBE = dr[2].ToString(); //DEBE
                    INTERES = dr[3].ToString();//SUBTOTAL DE LA NOTA DE PAGO
                    IVA = dr[4].ToString(); //IVA DE LA NOTA DE PAGO
                    TOTALFACT = dr[5].ToString(); //TOTAL DE LA NOTA DE PAGO
                    STATUS = dr[6].ToString();//ESTATUS DE LA NOTA DE PAGO



                    if (STATUS == "CANCELADO")
                    {


                    }
                    else
                    {
                        //NUMERO DE LA CUENTA
                        conexionmy.Open();
                        SqlCommand comando_A = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta)VALUES('" + CUENTA_CAJA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                        comando_A.ExecuteNonQuery();
                        conexionmy.Close();
                        //LAS LEYENDAS
                        conexionmy.Open();
                        SqlCommand comando_B = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta)VALUES('" + FACTURA + "   " + "C-" + CONTRATO + "')", conexionmy);
                        comando_B.ExecuteNonQuery();
                        conexionmy.Close();
                        // EL MONTO DEL PRESTAMO
                        string monto_valor;

                        decimal unosidebed = decimal.Round(Convert.ToDecimal(TOTALFACT), 2, MidpointRounding.AwayFromZero);
                        monto_valor = string.Format("{0:0.000000}", unosidebed);


                        conexionmy.Open();
                        SqlCommand comando_D = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta)VALUES('" + monto_valor + ",1.00" + "')", conexionmy);
                        comando_D.ExecuteNonQuery();
                        conexionmy.Close();



                        conexionmy.Open();
                        SqlCommand comando_Ia = new SqlCommand("USE " + server + " INSERT into cobros_poliza_interes(cuenta) VALUES('" + CUENTA_INTERES + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                        comando_Ia.ExecuteNonQuery();
                        conexionmy.Close();

                        //ahora leyenda interes
                        conexionmy.Open();
                        SqlCommand comando_Iaa = new SqlCommand("USE " + server + " INSERT into cobros_poliza_interes(cuenta) VALUES('" + "INTERESES- " + FACTURA + " " + "C-" + CONTRATO + "')", conexionmy);
                        comando_Iaa.ExecuteNonQuery();
                        conexionmy.Close();


                        //un nuevo espacio
                        conexionmy.Open();
                        SqlCommand comando_Iaabk = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta) VALUES('')", conexionmy);
                        comando_Iaabk.ExecuteNonQuery();
                        conexionmy.Close();

                        string valor_convertido3;
                        //h13 = Convert.ToString(INTERES);
                        //valor_convertido3 = string.Format(h13, "####0.000000");

                        decimal unosidebedep = decimal.Round(Convert.ToDecimal(INTERES), 2, MidpointRounding.AwayFromZero);
                        valor_convertido3 = string.Format("{0:0.000000}", unosidebedep);

                        conexionmy.Open();
                        SqlCommand comando_Iz = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta)VALUES('" + valor_convertido3 + ",1.00" + "')", conexionmy);
                        comando_Iz.ExecuteNonQuery();
                        conexionmy.Close();

                        //'ahora la cuenta del iva
                        conexionmy.Open();
                        SqlCommand comando_Iza = new SqlCommand("USE " + server + " INSERT into cobros_poliza_interes(cuenta) VALUES('" + CUENTA_IVA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                        comando_Iza.ExecuteNonQuery();
                        conexionmy.Close();
                        //leyenda iva
                        conexionmy.Open();
                        SqlCommand comando_Iaabe = new SqlCommand("USE " + server + " INSERT into cobros_poliza_interes(cuenta) VALUES('" + "IVA- " + FACTURA + "   " + "C-" + CONTRATO + "')", conexionmy);
                        comando_Iaabe.ExecuteNonQuery();
                        conexionmy.Close();
                        //un nuevo espacio
                        conexionmy.Open();
                        SqlCommand comando_Iaaba = new SqlCommand("USE " + server + " INSERT into cobros_poliza_interes(cuenta)VALUES('')", conexionmy);
                        comando_Iaaba.ExecuteNonQuery();
                        conexionmy.Close();

                        string valor_convertido35;
                        //h135 = Convert.ToString(IVA);
                        //valor_convertido35 = string.Format(h135, "####0.000000");

                        decimal unosidebedepa = decimal.Round(Convert.ToDecimal(IVA), 2, MidpointRounding.AwayFromZero);
                        valor_convertido35 = string.Format("{0:0.000000}", unosidebedepa);

                        conexionmy.Open();
                        SqlCommand comando_IzZ = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta)VALUES('" + valor_convertido35 + ",1.00" + "')", conexionmy);
                        comando_IzZ.ExecuteNonQuery();
                        conexionmy.Close();


                    }

                    #region TRASH
                    //ahora busco le fecha de origen del contrato en cuestion

                    // DataTable fecha_con = new DataTable();
                    // fecha_con.Clear();
                    // SqlDataAdapter fecha_Carga = new SqlDataAdapter("USE " + server +
                    //" SELECT fechacons, prestamo FROM  contratos where contrato='" + CONTRATO + "' " +
                    //"", conexionmy);
                    // fecha_Carga.Fill(fecha_con);
                    // // this.dataGridView2.DataSource = fecha_con;
                    // if (fecha_con.Rows.Count == 0)
                    // {
                    //     //exportar con nota de no hubo operacion este dia y seguir con rutina
                    // }
                    // else
                    // {
                    //     FECHACONS = fecha_con.Rows[0][0].ToString(); //this.dataGridView2.Rows[0].Cells[0].Value.ToString();//fechacons
                    //     PRESTAMO = fecha_con.Rows[0][1].ToString();//this.dataGridView2.Rows[0].Cells[1].Value.ToString();//prestamo

                    //     int año_inicial_uno, nom_uno, valor_uno;
                    //     string seis_uno = "";
                    //     string cuatro_uno;

                    //     string año_fechacons, mes_fechacons;
                    //     año_fechacons = Convert.ToDateTime(FECHACONS).ToString("yyyy");
                    //     mes_fechacons = Convert.ToDateTime(FECHACONS).ToString("MM");


                    //     año_inicial_uno = 2006;
                    //     nom_uno = 1;
                    //     valor_uno = Convert.ToInt32(año_fechacons);//año actual

                    //     cuatro_uno = "0" + mes_fechacons;

                    //     while (año_inicial_uno < valor_uno)
                    //     {
                    //         nom_uno = nom_uno + 1;
                    //         año_inicial_uno = año_inicial_uno + 1;

                    //     }
                    //     if (nom_uno < 10)
                    //     {
                    //         seis_uno = Convert.ToString("00" + nom_uno);
                    //     }
                    //     if (nom_uno >= 10)
                    //     {
                    //         seis_uno = Convert.ToString("0" + nom_uno);
                    //     }

                    //     quince = CUENTA_PRESTAMOS + "-" + seis_uno + "-" + cuatro_uno;

                    // }

                    //

                    //AHORA LA CUENTA DE LA CAJA
                    //conexionmy.Open();
                    //SqlCommand comando_E = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + quince + "                ," + "  " + DEPARTAMENTO + " ')", conexionmy);
                    //comando_E.ExecuteNonQuery();
                    //conexionmy.Close();
                    ////ahora leyenda caja
                    //conexionmy.Open();
                    //SqlCommand comando_F = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + FACTURA + "   " + "C-" + CONTRATO + "')", conexionmy);
                    //comando_F.ExecuteNonQuery();
                    //conexionmy.Close();
                    ////inserto espacio
                    ////ahora la el monto de prestamo
                    //conexionmy.Open();
                    //SqlCommand comando_G = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('')", conexionmy);
                    //comando_G.ExecuteNonQuery();
                    //conexionmy.Close();
                    ////nuevamene la cantidad
                    ////ahora la el monto de prestamo             

                    //string valor_convertido;


                    //decimal unosidebede = decimal.Round(Convert.ToDecimal(PRESTAMO), 2, MidpointRounding.AwayFromZero);
                    //valor_convertido = string.Format("{0:0.000000}", unosidebede);

                    //if (STATUS == "CANCELADO")
                    //{
                    //    conexionmy.Open();
                    //    SqlCommand comando_H = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                    //    comando_H.ExecuteNonQuery();
                    //    conexionmy.Close();
                    //}
                    //else
                    //{
                    //    conexionmy.Open();
                    //    SqlCommand comando_I = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + valor_convertido + ",1.00" + "')", conexionmy);
                    //    comando_I.ExecuteNonQuery();
                    //    conexionmy.Close();
                    //}
                    //'ahora cuenta caja
                    //conexionmy.Open();
                    //SqlCommand comando_Ia = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + CUENTA_INTERES + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    //comando_Ia.ExecuteNonQuery();
                    //conexionmy.Close();

                    ////ahora leyenda interes
                    //conexionmy.Open();
                    //SqlCommand comando_Iaa = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + "INTERESES- " + FACTURA + " " + "C-" + CONTRATO + "')", conexionmy);
                    //comando_Iaa.ExecuteNonQuery();
                    //conexionmy.Close();


                    ////un nuevo espacio
                    //conexionmy.Open();
                    //SqlCommand comando_Iaabk = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta) VALUES('')", conexionmy);
                    //comando_Iaabk.ExecuteNonQuery();
                    //conexionmy.Close();
                    //ingreso del interes    
                    //string h13;
                    //string valor_convertido3;
                    ////h13 = Convert.ToString(INTERES);
                    ////valor_convertido3 = string.Format(h13, "####0.000000");

                    //decimal unosidebedep = decimal.Round(Convert.ToDecimal(INTERES), 2, MidpointRounding.AwayFromZero);
                    //valor_convertido3 = string.Format("{0:0.000000}", unosidebedep);

                    //if (STATUS == "CANCELADO")
                    //{
                    //    conexionmy.Open();
                    //    SqlCommand comando_Ha = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                    //    comando_Ha.ExecuteNonQuery();
                    //    conexionmy.Close();
                    //}
                    //else
                    //{
                    //    conexionmy.Open();
                    //    SqlCommand comando_Iz = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + valor_convertido3 + ",1.00" + "')", conexionmy);
                    //    comando_Iz.ExecuteNonQuery();
                    //    conexionmy.Close();
                    //}

                    ////'ahora la cuenta del iva
                    //conexionmy.Open();
                    //SqlCommand comando_Iza = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + CUENTA_IVA + "                " + ",  " + DEPARTAMENTO + "')", conexionmy);
                    //comando_Iza.ExecuteNonQuery();
                    //conexionmy.Close();
                    ////leyenda iva
                    //conexionmy.Open();
                    //SqlCommand comando_Iaabe = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta) VALUES('" + "IVA- " + FACTURA + "   " + "C-" + CONTRATO + "')", conexionmy);
                    //comando_Iaabe.ExecuteNonQuery();
                    //conexionmy.Close();
                    ////un nuevo espacio
                    //conexionmy.Open();
                    //SqlCommand comando_Iaaba = new SqlCommand("USE " + server + " INSERT into cobros_poliza(cuenta)VALUES('')", conexionmy);
                    //comando_Iaaba.ExecuteNonQuery();
                    //conexionmy.Close();

                    //'valor del iva
                    //string h135;
                    //string valor_convertido35;
                    ////h135 = Convert.ToString(IVA);
                    ////valor_convertido35 = string.Format(h135, "####0.000000");

                    //decimal unosidebedepa = decimal.Round(Convert.ToDecimal(IVA), 2, MidpointRounding.AwayFromZero);
                    //valor_convertido35 = string.Format("{0:0.000000}", unosidebedepa);

                    //if (STATUS == "CANCELADO")
                    //{
                    //    conexionmy.Open();
                    //    SqlCommand comando_Ha = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + "0.01" + ",1.00" + "')", conexionmy);
                    //    comando_Ha.ExecuteNonQuery();
                    //    conexionmy.Close();
                    //}
                    //else
                    //{
                    //    conexionmy.Open();
                    //    SqlCommand comando_Iz = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza(cuenta)VALUES('" + valor_convertido35 + ",1.00" + "')", conexionmy);
                    //    comando_Iz.ExecuteNonQuery();
                    //    conexionmy.Close();
                    //}

                    // 
                    #endregion
                }//cierre del ciclo for each
                conexionmy.Open();
                SqlCommand comando_J = new SqlCommand("USE " + server + " INSERT INTO cobros_poliza_interes(cuenta)VALUES('FIN')", conexionmy);
                comando_J.ExecuteNonQuery();
                conexionmy.Close();
                bloc_SoloInteresEivaCaja();
            }


        }
        //CREAR ARCHIVO POR ARCHVIVO POOL DE INTERES
        public void bloc_SoloInteresEivaCaja()
        {
            SqlConnection conexionmy = new SqlConnection(sqlcnx);
            //conexionmy.ConnectionString = sqlcnx;//conexion mysql
            DataTable tablaprestamos = new DataTable();
            tablaprestamos.Clear();
            SqlDataAdapter datos_presa = new SqlDataAdapter("USE " + server + "  " +
                "SELECT * FROM  cobros_poliza_interes " +
                " order by  no asc " +
                "", conexionmy);
            datos_presa.Fill(tablaprestamos);



            if (!Directory.Exists(path + "/INTERES-" + caja))
            {
                Directory.CreateDirectory(path + "/INTERES-" + caja);
            }
            System.IO.StreamWriter escribir = new System.IO.StreamWriter(path + "/INTERES-" + caja + "/INTERES " + caja + "-" + fecha_uno + ".pol");
            foreach (DataRow fila in tablaprestamos.Rows)
            {
                escribir.WriteLine(fila[1]);
            }
            escribir.Close();

        }

        #endregion


        #region BackGroundWorkers

        #region identificaServidor
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
        #endregion

        #region PrestamosB 
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



        }
        #endregion

        #region CobrosB
        private void backgroundWorker3_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            nveces2 = 0;


        }

        private void backgroundWorker3_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBarX2.Value = e.ProgressPercentage;
        }

        private void backgroundWorker3_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            cobros();
        }
        #endregion

        #region Interes DiarioB
        private void backgroundWorker4_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            nveces3 = 0;



        }

        private void backgroundWorker4_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBarX3.Value = e.ProgressPercentage;
        }

        private void backgroundWorker4_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            diario_interes();
        }

        #endregion

        #region Interes SemanalB
        private void backgroundWorker5_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            poliza_interes_semanal();
        }

        private void backgroundWorker5_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBarX4.Value = e.ProgressPercentage;
        }

        private void backgroundWorker5_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {

        }

        #endregion

        #region Remision SemanalB
        private void backgroundWorker6_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {


        }

        private void comboBoxEx3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBoxEx2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }



        private void backgroundWorker6_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBarX5.Value = e.ProgressPercentage;
        }

        private void backgroundWorker6_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            poliza_remisiones_semanal();
        }
        #endregion


        #region ContratosDesempeñados
        private void backgroundWorker7_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            ContratosDesempeñados();
        }



        private void backgroundWorker7_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBarX6.Value = e.ProgressPercentage;
        }

        private void backgroundWorker7_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {

        }
        #endregion

        #region Interes e IVA por caja
        private void backgroundWorker8_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            SoloInteresEivaCaja();
        }

        private void checkBoxX6_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBoxX7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtDias_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = false;
            }
            else
           if (Char.IsControl(e.KeyChar)) //permitir teclas de control como retroceso
            {
                e.Handled = false;
            }
            else
            {
                //el resto de teclas pulsadas se desactivan
                e.Handled = true;
            }
        }

        private void txtDiasFinal_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = false;
            }
            else
            if (Char.IsControl(e.KeyChar)) //permitir teclas de control como retroceso
            {
                e.Handled = false;
            }
            else
            {
                //el resto de teclas pulsadas se desactivan
                e.Handled = true;
            }
        }

        private void backgroundWorker8_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBarX7.Value = e.ProgressPercentage;
        }

        private void backgroundWorker8_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {

        }

        #endregion




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


        #endregion


    }
}
