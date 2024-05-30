﻿using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using OfficeOpenXml;
using Xceed.Words.NET;
using System.ComponentModel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Xml.Serialization;
using Newtonsoft.Json;  // Asegúrate de que este using esté presente


namespace WinFormsApp6
{
    public partial class Form1 : Form
    {

        List<Persona> personas = new List<Persona>();
        private DataGridView dgvDatos;
        private DateTimePicker dtpFechaNacimiento;
        private DateTimePicker dtpFechaRegistro;
        private TextBox txtNombre;
        private TextBox txtApellido;
        private ErrorProvider errorProvider1;
        private IContainer components;
        private Button btnAgregar;
        private Button btnGuardarCSV;
        private Button btnMostrarData;
        private Button btnPDF;
        private Button btnExcel;
        private Button btnWord;
        private Button btnJSON;
        private Button btnXML;
        private Label label1;
        private Label label4;
        private Label label3;
        private Label label2;
        private Label label5;
        string csvFilePath = "personas.csv"; // Nombre del archivo CSV


        public Form1()
        {

            InitializeComponent();
            CargarDatosDesdeCSV(); // Cargar los datos desde el archivo CSV al iniciar la aplicación


        }

        private void ExportarDatosAExcel()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel files (*.xlsx)|*.xlsx";
                sfd.Title = "Guardar como Excel";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    using (ExcelPackage excel = new ExcelPackage())
                    {
                        var worksheet = excel.Workbook.Worksheets.Add("Personas");
                        var headerRow = new List<string[]> { new string[] { "ID", "Nombres", "Apellidos", "Fecha Nacimiento", "Fecha Registro" } };

                        // Determinamos las propiedades del encabezado
                        worksheet.Cells["A1:E1"].LoadFromArrays(headerRow);

                        // Añadimos los datos
                        int rowIndex = 2;
                        foreach (var persona in personas)
                        {
                            worksheet.Cells[rowIndex, 1].Value = persona.Id;
                            worksheet.Cells[rowIndex, 2].Value = persona.Nombres;
                            worksheet.Cells[rowIndex, 3].Value = persona.Apellidos;
                            worksheet.Cells[rowIndex, 4].Value = persona.FechaNcimiento.ToString("yyyy-MM-dd");
                            worksheet.Cells[rowIndex, 5].Value = persona.FechaRegistro.ToString("yyyy-MM-dd");
                            rowIndex++;
                        }

                        // Guardamos el archivo
                        var fileInfo = new FileInfo(sfd.FileName);
                        excel.SaveAs(fileInfo);
                        MessageBox.Show("Datos exportados a Excel correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void ExportarDatosAWord()
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Word files (*.docx)|*.docx";
                sfd.Title = "Guardar como Word";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    using (var doc = DocX.Create(sfd.FileName))
                    {
                        doc.InsertParagraph("Lista de Personas").FontSize(20).Bold();
                        foreach (var persona in personas)
                        {
                            doc.InsertParagraph($"{persona.Id}, {persona.Nombres}, {persona.Apellidos}, {persona.FechaNcimiento:yyyy-MM-dd}, {persona.FechaRegistro:yyyy-MM-dd}");
                        }
                        doc.Save();
                        MessageBox.Show("Datos exportados a Word correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }
        private void ExportarDatosAPDF()
        {

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "PDF files (*.pdf)|*.pdf";
                sfd.Title = "Guardar como PDF";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        Document doc = new Document();
                        PdfWriter writer = PdfWriter.GetInstance(doc, fs);
                        doc.Open();

                        foreach (var persona in personas)
                        {
                            doc.Add(new Paragraph($"{persona.Id}, {persona.Nombres}, {persona.Apellidos}, {persona.FechaNcimiento:yyyy-MM-dd}, {persona.FechaRegistro:yyyy-MM-dd}"));
                        }

                        doc.Close();
                        writer.Close();
                        MessageBox.Show("Datos exportados a PDF correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }


        private void ExportarDatosAXML()
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "XML files (*.xml)|*.xml";
                sfd.Title = "Guardar como XML";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(List<Persona>));
                    using (StreamWriter sw = new StreamWriter(sfd.FileName))
                    {
                        serializer.Serialize(sw, personas);
                    }
                    MessageBox.Show("Datos exportados a XML correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        private void ExportarDatosAJSON()
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "JSON files (*.json)|*.json";
                sfd.Title = "Guardar como JSON";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string json = JsonConvert.SerializeObject(personas, Formatting.Indented);
                    File.WriteAllText(sfd.FileName, json);
                    MessageBox.Show("Datos exportados a JSON correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void CargarDatosDesdeCSV()
        {
            try
            {
                if (File.Exists(csvFilePath))
                {
                    using (var reader = new StreamReader(csvFilePath))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            var values = line.Split(',');

                            var persona = new Persona
                            {
                                Id = int.Parse(values[0]),
                                Nombres = values[1],
                                Apellidos = values[2],
                                FechaNcimiento = DateTime.Parse(values[3]),
                                FechaRegistro = DateTime.Parse(values[4])
                            };

                            personas.Add(persona);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar los datos: " + ex.Message);
            }

            dgvDatos.DataSource = personas;
        }





        private void InitializeComponent()

        {
            this.components = new System.ComponentModel.Container();
            this.dgvDatos = new System.Windows.Forms.DataGridView();
            this.dtpFechaNacimiento = new System.Windows.Forms.DateTimePicker();
            this.dtpFechaRegistro = new System.Windows.Forms.DateTimePicker();
            this.txtNombre = new System.Windows.Forms.TextBox();
            this.txtApellido = new System.Windows.Forms.TextBox();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.btnAgregar = new System.Windows.Forms.Button();
            this.btnGuardarCSV = new System.Windows.Forms.Button();
            this.btnMostrarData = new System.Windows.Forms.Button();
            this.btnWord = new System.Windows.Forms.Button();
            this.btnExcel = new System.Windows.Forms.Button();
            this.btnPDF = new System.Windows.Forms.Button();
            this.btnXML = new System.Windows.Forms.Button();
            this.btnJSON = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDatos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvDatos
            // 
            this.dgvDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDatos.Location = new System.Drawing.Point(318, 50);
            this.dgvDatos.Name = "dgvDatos";
            this.dgvDatos.RowTemplate.Height = 25;
            this.dgvDatos.Size = new System.Drawing.Size(390, 274);
            this.dgvDatos.TabIndex = 0;
            // 
            // dtpFechaNacimiento
            // 
            this.dtpFechaNacimiento.Location = new System.Drawing.Point(12, 62);
            this.dtpFechaNacimiento.Name = "dtpFechaNacimiento";
            this.dtpFechaNacimiento.Size = new System.Drawing.Size(109, 23);
            this.dtpFechaNacimiento.TabIndex = 1;
            // 
            // dtpFechaRegistro
            // 
            this.dtpFechaRegistro.Location = new System.Drawing.Point(185, 65);
            this.dtpFechaRegistro.Name = "dtpFechaRegistro";
            this.dtpFechaRegistro.Size = new System.Drawing.Size(109, 23);
            this.dtpFechaRegistro.TabIndex = 2;
            // 
            // txtNombre
            // 
            this.txtNombre.Location = new System.Drawing.Point(12, 21);
            this.txtNombre.Name = "txtNombre";
            this.txtNombre.Size = new System.Drawing.Size(167, 23);
            this.txtNombre.TabIndex = 3;
            // 
            // txtApellido
            // 
            this.txtApellido.Location = new System.Drawing.Point(185, 21);
            this.txtApellido.Name = "txtApellido";
            this.txtApellido.Size = new System.Drawing.Size(167, 23);
            this.txtApellido.TabIndex = 4;
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // btnAgregar
            // 
            this.btnAgregar.Location = new System.Drawing.Point(12, 91);
            this.btnAgregar.Name = "btnAgregar";
            this.btnAgregar.Size = new System.Drawing.Size(148, 46);
            this.btnAgregar.TabIndex = 5;
            this.btnAgregar.Text = "Agregar a dgv";
            this.btnAgregar.UseVisualStyleBackColor = true;
            this.btnAgregar.Click += new System.EventHandler(this.btnAgregar_Click_1);
            // 
            // btnGuardarCSV
            // 
            this.btnGuardarCSV.Location = new System.Drawing.Point(166, 94);
            this.btnGuardarCSV.Name = "btnGuardarCSV";
            this.btnGuardarCSV.Size = new System.Drawing.Size(140, 45);
            this.btnGuardarCSV.TabIndex = 6;
            this.btnGuardarCSV.Text = "Guardar en CSV";
            this.btnGuardarCSV.UseVisualStyleBackColor = true;
            this.btnGuardarCSV.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnMostrarData
            // 
            this.btnMostrarData.Location = new System.Drawing.Point(14, 145);
            this.btnMostrarData.Name = "btnMostrarData";
            this.btnMostrarData.Size = new System.Drawing.Size(232, 22);
            this.btnMostrarData.TabIndex = 7;
            this.btnMostrarData.Text = "Mostrar Data";
            this.btnMostrarData.UseVisualStyleBackColor = true;
            this.btnMostrarData.Click += new System.EventHandler(this.btnMostrarData_Click);
            // 
            // btnWord
            // 
            this.btnWord.Location = new System.Drawing.Point(264, 190);
            this.btnWord.Name = "btnWord";
            this.btnWord.Size = new System.Drawing.Size(48, 22);
            this.btnWord.TabIndex = 8;
            this.btnWord.Text = "Word";
            this.btnWord.UseVisualStyleBackColor = true;
            this.btnWord.Click += new System.EventHandler(this.btnWord_Click);
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(264, 218);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(48, 22);
            this.btnExcel.TabIndex = 9;
            this.btnExcel.Text = "Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // btnPDF
            // 
            this.btnPDF.Location = new System.Drawing.Point(264, 246);
            this.btnPDF.Name = "btnPDF";
            this.btnPDF.Size = new System.Drawing.Size(48, 22);
            this.btnPDF.TabIndex = 10;
            this.btnPDF.Text = "PDF";
            this.btnPDF.UseVisualStyleBackColor = true;
            this.btnPDF.Click += new System.EventHandler(this.btnPDF_Click);
            // 
            // btnXML
            // 
            this.btnXML.Location = new System.Drawing.Point(264, 274);
            this.btnXML.Name = "btnXML";
            this.btnXML.Size = new System.Drawing.Size(48, 22);
            this.btnXML.TabIndex = 11;
            this.btnXML.Text = "XML";
            this.btnXML.UseVisualStyleBackColor = true;
            this.btnXML.Click += new System.EventHandler(this.btnXML_Click);
            // 
            // btnJSON
            // 
            this.btnJSON.Location = new System.Drawing.Point(264, 302);
            this.btnJSON.Name = "btnJSON";
            this.btnJSON.Size = new System.Drawing.Size(48, 22);
            this.btnJSON.TabIndex = 12;
            this.btnJSON.Text = "JSON";
            this.btnJSON.UseVisualStyleBackColor = true;
            this.btnJSON.Click += new System.EventHandler(this.btnJSON_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 15);
            this.label1.TabIndex = 13;
            this.label1.Text = "Nombre";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(195, 3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 15);
            this.label2.TabIndex = 14;
            this.label2.Text = "Apellido";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 47);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(120, 15);
            this.label3.TabIndex = 15;
            this.label3.Text = "Fecha De Nacimiento";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(185, 47);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(101, 15);
            this.label4.TabIndex = 16;
            this.label4.Text = "Fecha De Registro";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(264, 172);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(50, 15);
            this.label5.TabIndex = 17;
            this.label5.Text = "Exportar";
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(792, 374);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnJSON);
            this.Controls.Add(this.btnXML);
            this.Controls.Add(this.btnPDF);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnWord);
            this.Controls.Add(this.btnMostrarData);
            this.Controls.Add(this.btnGuardarCSV);
            this.Controls.Add(this.btnAgregar);
            this.Controls.Add(this.txtApellido);
            this.Controls.Add(this.txtNombre);
            this.Controls.Add(this.dtpFechaRegistro);
            this.Controls.Add(this.dtpFechaNacimiento);
            this.Controls.Add(this.dgvDatos);
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDatos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

       
        private void btnAgregar_Click_1(object sender, EventArgs e)
        {
            if (txtNombre.Text == "")
            {
                errorProvider1.SetError(txtNombre, "Debe ingresar su Nombre");
                txtNombre.Focus();
                return;
            }
            errorProvider1.SetError(txtNombre, "");

            if (txtApellido.Text == "")
            {
                errorProvider1.SetError(txtApellido, "Debe ingresar su Apellido");
                txtApellido.Focus();
                return;
            }
            errorProvider1.SetError(txtApellido, "");

            Persona mipersona = new Persona(txtNombre.Text, txtApellido.Text, dtpFechaNacimiento.Value);
            personas.Add(mipersona);

            dgvDatos.DataSource = null;
            dgvDatos.DataSource = personas;

            txtNombre.Clear();
            txtApellido.Clear();
            txtNombre.Focus();

            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSV files (*.csv)|*.csv";
                sfd.Title = "Guardar como CSV";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (StreamWriter sw = new StreamWriter(sfd.FileName))
                        {
                            // Escribir los encabezados del DataGridView
                            for (int i = 0; i < dgvDatos.Columns.Count; i++)
                            {
                                sw.Write(dgvDatos.Columns[i].HeaderText);
                                if (i < dgvDatos.Columns.Count - 1)
                                {
                                    sw.Write(",");
                                }
                            }
                            sw.WriteLine();

                            // Escribir los datos del DataGridView
                            foreach (DataGridViewRow row in dgvDatos.Rows)
                            {
                                if (!row.IsNewRow)
                                {
                                    for (int i = 0; i < dgvDatos.Columns.Count; i++)
                                    {
                                        sw.Write(row.Cells[i].Value?.ToString());
                                        if (i < dgvDatos.Columns.Count - 1)
                                        {
                                            sw.Write(",");
                                        }
                                    }
                                    sw.WriteLine();
                                }
                            }
                        }
                        MessageBox.Show("Datos guardados correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al guardar los datos: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnMostrarData_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "CSV files (*.csv)|*.csv";
                ofd.Title = "Abrir archivo CSV";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        List<Persona> nuevasPersonas = new List<Persona>();
                        using (StreamReader sr = new StreamReader(ofd.FileName))
                        {
                            // Leer los encabezados
                            string[] headers = sr.ReadLine().Split(',');

                            // Leer las filas
                            while (!sr.EndOfStream)
                            {
                                string[] fields = sr.ReadLine().Split(',');
                                Persona personas = new Persona()
                                {
                                    Id = int.Parse(fields[0]),
                                    Nombres = fields[1],
                                    Apellidos = fields[2],
                                    FechaNcimiento = DateTime.Parse(fields[3]),
                                    FechaRegistro = DateTime.Parse(fields[4])
                                };
                                nuevasPersonas.Add(personas);
                            }
                        }

                        // Actualizar la lista de personas y el DataGridView
                        personas = nuevasPersonas;
                        dgvDatos.DataSource = null;
                        dgvDatos.DataSource = personas;

                        MessageBox.Show("Datos cargados correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al cargar los datos: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnWord_Click(object sender, EventArgs e)
        {

            ExportarDatosAWord();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            ExportarDatosAExcel();
        }

        private void btnPDF_Click(object sender, EventArgs e)
        { 
            ExportarDatosAPDF();
                
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnXML_Click(object sender, EventArgs e)
        {
            ExportarDatosAXML();
        }

        private void btnJSON_Click(object sender, EventArgs e)
        {
            ExportarDatosAJSON();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}


 
 
    