using Entities;
using Microsoft.VisualBasic.Devices;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Dev_Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //Instance of Folder_Files class
        Folder_Files _Files = new Folder_Files();
        List<Folder_Files> list_files = new List<Folder_Files>();
        List<Folder_Files> list_files_2 = new List<Folder_Files>();

        string folder = "";
        string[] files = { };
        int count_Excel = 1;

        Actions actions = new Actions();
        Computer myComputer = new Computer();

        #region Abrir carpeta para examinar Ruta
        private void button1_Click(object sender, EventArgs e)
        {
            new_folder();
            change_dataGrid();
            create_folder();
            merge(files, count_Excel);
        }

        private void create_folder()
        {
            string move_to = "";

            move_to = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            //using (var fd = new FolderBrowserDialog())
            //{
            //    if (fd.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fd.SelectedPath))
            //    {
            //        folder_name.Text = fd.SelectedPath.Substring(fd.SelectedPath.LastIndexOf(@"\"));
            //        move_to = fd.SelectedPath;
            //    }

            //}



            string Processed = move_to + @"\Processed";
            string Not_Applicable = move_to + @"\Not Applicable";
            string Master_File = move_to + @"\Processed\Master File";


            if (!Directory.Exists(Processed))
            {
                Directory.CreateDirectory(Processed);
                Console.WriteLine(Processed);

                

            }
            if (!Directory.Exists(Master_File))
            {
                Directory.CreateDirectory(Master_File);
                Console.WriteLine(Master_File);



            }
            if (!Directory.Exists(Not_Applicable))
            {
                Directory.CreateDirectory(Not_Applicable);
                Console.WriteLine(Not_Applicable);



            }

            move_file(move_to);
        }

        private void new_folder()
        {
            using (var fd = new FolderBrowserDialog())
            {

                if (fd.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fd.SelectedPath))
                {
                    folder_name.Text = fd.SelectedPath.Substring(fd.SelectedPath.LastIndexOf(@"\"));
                    folder = fd.SelectedPath;
                }

            }

            add_route_to_Excel();

            
            //change_file(folder_name.Text);
        }
        #endregion

        #region Cargar los datos de los archivos y sus rutas en un Excel

        private void add_route_to_Excel()
        {

            string pathFile = AppDomain.CurrentDomain.BaseDirectory + "recent_Folder.xlsx";

            SLDocument oSLDocument = new SLDocument();

            DataTable dt = new DataTable();

            //columnas
            dt.Columns.Add("Numero", typeof(int));
            dt.Columns.Add("Nombre", typeof(string));
            dt.Columns.Add("Ruta", typeof(string));

            change_list(list_files, folder);

            //registros , rows
            foreach (Folder_Files file_save in list_files)
            {
                dt.Rows.Add(file_save.ID, file_save.name, file_save.folder);
            }

            oSLDocument.ImportDataTable(1, 1, dt, true);
            oSLDocument.SaveAs(pathFile);
        }



        #endregion

        #region Cargar la lista con las Rutas en la lista

        private void change_list(List<Folder_Files> list_files, string folder_name)
        {
            
            if(folder_name==string.Empty || folder_name == "")
            {
                folder_name = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }

            DirectoryInfo di = new DirectoryInfo(folder_name);
            DirectoryInfo di1 = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Processed");
            
            FileInfo[] files = di.GetFiles();
            FileInfo[] files1 = di1.GetFiles();

            FileInfo fInfo = new FileInfo(folder_name);

            int cont = 0, activo = 1;

            list_files.Clear();

            foreach (FileInfo file in files)
            {
                Folder_Files _files = new Folder_Files();

                if (file.Name.EndsWith("xls") || file.Name.EndsWith("xlsx"))
                {
                    _files.ID = activo;
                    _files.name = file.Name;
                    _files.folder = file.FullName;

                    cont++;
                    list_files.Add(_files);

                    activo++;
                }
                else if (!file.Name.EndsWith("mp4") || !file.Name.EndsWith("exe") || !file.Name.EndsWith("msi"))
                {
                    _files.ID = activo;
                    _files.name = file.Name;
                    _files.folder = file.FullName;

                    list_files.Add(_files);

                    activo++;

                }

                /*dg.Rows.Add();
                dg.Rows[cont].Cells[0].Value = activo;
                dg.Rows[cont].Cells[1].Value = file.Name;*/
            }

            list_files_2.Clear();
            foreach (FileInfo file in files1)
            {
                Folder_Files _files = new Folder_Files();
                if (file.Name.EndsWith("xls") || file.Name.EndsWith("xlsx"))
                {
                    _files.ID = activo;
                    _files.name = file.Name;
                    _files.folder = file.FullName;

                    cont++;
                    list_files_2.Add(_files);

                    activo++;
                }

            }
        }


        #endregion


        #region Crear las carpetas  y moverlos a sus carpetas
        private void button2_Click(object sender, EventArgs e) { }

        private void move_file(string route)
        {
            int xlsx =list_files_2.Count();
            foreach(Folder_Files _Files in list_files)
            {

                try
                {
                    if (_Files.name.EndsWith("xls") || _Files.name.EndsWith("xlsx"))
                    {
                        myComputer.FileSystem.CopyFile(_Files.folder, route + @"\Processed\" + _Files.name);

                    }
                    else if(!_Files.name.EndsWith("mp4") || !_Files.name.EndsWith("exe") || !_Files.name.EndsWith("msi"))
                    {
                        myComputer.FileSystem.CopyFile(_Files.folder, route + @"\Not Applicable\" + _Files.name);

                    }

                }
                catch(Exception ex)
                {

                    //MessageBox.Show("El archivo ya existe");
                }
            }


        }

        

        #endregion

        private void button3_Click(object sender, EventArgs e)
        {

            
        }

        private void merge(string[] files, int count_Excel)
        {
            //list_files_2.Select(x => x.folder).ToArray();
            string route = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Processed\Master File\Master File";

            //folder_name = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            actions.MergeXlsxFiles(route, list_files_2.Select(x => x.folder).ToArray());


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            change_dataGrid();

        }

        private void change_dataGrid()
        {
            change_list(list_files, folder);

            int count = 0;
            dataGridView1.Rows.Clear();
            foreach (Folder_Files folder_Files in list_files)
            {
                dataGridView1.Rows.Add();

                dataGridView1.Rows[count].Cells[0].Value = folder_Files.ID;
                dataGridView1.Rows[count].Cells[1].Value = folder_Files.name;
                dataGridView1.Rows[count].Cells[2].Value = folder_Files.folder;

                count++;

            }
        }
    }

}
