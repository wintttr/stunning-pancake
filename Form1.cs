using System.Windows.Forms;
using System.Threading;
using System.ComponentModel;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        private  TeachPlan? teachPlan = null;

        public Form1()
        {
            InitializeComponent();
        }
        private void loadFileButton_Click(object sender, EventArgs e)
        {
            using OpenFileDialog ofd = new();

            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.DefaultExt = "xlsx";
            ofd.Filter = "Excel files (*.xlsx)|*.xlsx|All files(*.*)|*.* ";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    ExcelReader er = new(ofd.FileName);
                    teachPlan = er.GetPlan();

                    disciplineCheckedList.Items.Clear();
                    foreach (var d in teachPlan.Disciplines)
                        disciplineCheckedList.Items.Add(d);
                }
                catch (ParseErrorException ex)
                {
                    MessageBox.Show($"Произошла ошибка при разборе файла: {ex.Message}");
                }
            }
        }

        private void generateOutputButton_Click(object sender, EventArgs e)
        {
            if (teachPlan is null)
            {
                MessageBox.Show("Файл не загружен.");
                return;
            }

            using FolderBrowserDialog fbd = new();

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                PlaceHolder ph = new PlaceHolder(teachPlan, fbd.SelectedPath);
            }
        }
    }
}