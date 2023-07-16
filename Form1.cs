using System.Windows.Forms;
using System.Threading;
using System.ComponentModel;
using System.Data;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        private TeachPlan? teachPlan = null;
        private ICollection<Discipline>? disciplines = null;
        private IDictionary<string, string>? compDisc = null;

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
                    teachPlan = er.Plan;
                    disciplines = er.Disciplines;
                    compDisc = er.CompDisc;

                    disciplineCheckedList.Items.Clear();
                    foreach (var d in disciplines)
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
            if (teachPlan is null || disciplines is null || compDisc is null)
            {
                MessageBox.Show("Файл не загружен.");
                return;
            }

            if (disciplineCheckedList.CheckedItems.Count == 0)
            {
                MessageBox.Show("Выберите хотя бы одну дисциплину.");
                return;
            }

            using FolderBrowserDialog fbd = new();

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                PlaceHolder ph = new PlaceHolder(teachPlan, compDisc, fbd.SelectedPath);

                foreach(Discipline d in disciplineCheckedList.CheckedItems)
                {
                    ph.OneDiscipline(d);
                }
            }

            MessageBox.Show("Сгенерировано.");
        }
    }
}