using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace WorkViewSQL
{

    public partial class frmWorkViewSQL : Form
    {

        System.Timers.Timer tmr = new System.Timers.Timer(1000);
        delegate void myDeleg(Form toClose);
        Form currentPrgBar = null;
        Boolean createPrgBar = false;
        Boolean disposePrgBar = false;

        public frmWorkViewSQL()
        {
            InitializeComponent();
        }

        private void btnImportData_Click(object sender, EventArgs e)
        {
            dataGridView1.Visible = false;
            Application.DoEvents();
            this.Cursor = Cursors.WaitCursor;
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            //dataGridView1.Refresh();
            simulateLoad();
            //Obtenemos el archivo desde la ubicación actual
            var executableFolderPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            //Hoja desde donde obtendremos los datos
            string hoja = "Hoja1";
            //Cadena de conexión
            string conexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + //executableFolderPath +
                            lblPath.Text +  //"C:\\paso\\SAP.xlsx" +
                            ";Extended Properties='Excel 8.0;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(conexion);
            //Consulta contra la hoja de Excel
            OleDbCommand cmd = new OleDbCommand("Select * From [" + hoja + "$]", con);
            try
            {
                //Conectarse al archivo de Excel
                con.Open();

                OleDbDataAdapter sda = new OleDbDataAdapter(cmd);
                DataTable data = new DataTable();

                //Cargar los datos
                sda.Fill(data);

                //Cargar la grilla
                dataGridView1.DataSource = data;
                //dataGridView1.AutoResizeColumns();
                //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView1.Visible = true;
                
            }
            catch
            {
                //Error leyendo excel
                MessageBox.Show("Ocurrió un error en la lectura del archivo");
                dataGridView1.Visible = false;
            }
            finally
            {
                //Funcione o no, cerramos la cadena de conexión
                con.Close();
            }
            this.Cursor = Cursors.Default;
        }


        public void simulateLoad()
        {
            createPrgBar = true;
            //btnRunning.Text = "Load in progress";
            for (int i = 0; i < 10000; i++)
            {
                //btnRunning.Text = "Running loop :" + i.ToString();
                //btnRunning.Refresh();
                Application.DoEvents();
            }
            //btnRunning.Text = "Run New Loop";
            disposePrgBar = true;
        }

        /// watch if any process are running to create a new progress bar
        void tmr_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (createPrgBar & (currentPrgBar == null))
            {
                Thread myThread = new Thread(onFlyProgressBar);
                myThread.Start();
                createPrgBar = false;
            }
            else
            {
                if (disposePrgBar)
                {
                    closeForm(currentPrgBar);
                    disposePrgBar = false;
                }
            }
        }

        /// Create new 'processing' progressBar
        void onFlyProgressBar()
        {
            Label lblMessage = new Label
            {
                Dock = DockStyle.Top,
                Text = "Procesing, please wait..",
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 14,
                                                System.Drawing.FontStyle.Bold,
                                                System.Drawing.GraphicsUnit.Point)
            };

            ProgressBar prgBar = new ProgressBar
            {
                Height = 15,
                Dock = DockStyle.Bottom,
                Style = ProgressBarStyle.Marquee
            };

            currentPrgBar = new Form()
            {
                Width = 300,
                Height = 50,
                StartPosition = FormStartPosition.CenterScreen,
                ControlBox = false,
                FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
            };
            currentPrgBar.Controls.AddRange(new Control[] { lblMessage, prgBar });
            currentPrgBar.LostFocus += recoverFocus;
            Application.Run(currentPrgBar);
        }

        void recoverFocus(object sender, EventArgs e)
        {
            currentPrgBar.TopMost = true;
        }

        /// Make current prgBar topmost if lost focus
        void closeForm(Form toClose)
        {
            if (toClose == null) return;
            if (toClose.InvokeRequired)
            {
                toClose.Invoke(new myDeleg(closeForm), toClose);
            }
            else
            {
                toClose.Close();
                currentPrgBar = null;
            }
        }


        private void frmWorkViewSQL_Load(object sender, EventArgs e)
        {
            tmr.Elapsed += tmr_Elapsed;
            tmr.Interval = 100;
            tmr.Start();

            lblPath.BackColor = System.Drawing.Color.Transparent;
            btnImportData.BackColor = System.Drawing.Color.Transparent;
        }

        private void lblPath_DoubleClick(object sender, EventArgs e)
        {
            ofd.Title = "Browse Text Files";
            ofd.Filter = "Excel files (*.xl*)|*.xl*|Text files (*.txt;*.csv)|*.txt;*.csv|XML files (*.xml)|*.xml|All files (*.*)|*.*";
            ofd.FilterIndex = 4;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.RestoreDirectory = true;
            ofd.InitialDirectory = @"C:\Station VY Canis Majoris\Proceso de Cierre\Clasificaciones ASTI_MAESTRO\Clasificacion Exp_SAP";  //Application.CommonAppDataPath; //@"C:\"
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                lblPath.Text = ofd.FileName;
            }
        }

        private void lblPath_DragOver(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length != 0)
            {
                lblPath.Text = files[0];
            }
        }

        private void frmWorkViewSQL_DragOver(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length != 0)
            {
                lblPath.Text = files[0];
            }
        }

        
    }
}
