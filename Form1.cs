using System.Windows.Forms;
using System.Threading;
using System.ComponentModel;
using System.Data;
using System.Security.Cryptography;

namespace WinFormsApp1
{
    public partial class StunningPancakeForm : Form
    {
        private TeachPlan? teachPlan = null;
        private ICollection<Discipline>? disciplines = null;
        private IDictionary<string, string>? compDisc = null;

        BackgroundWorker saveFileWorker = new();
        BackgroundWorker loadFileWorker = new();

        public StunningPancakeForm()
        {
            InitializeComponent();

            loadFileWorker.DoWork += LoadFileWorker_DoWork;
            loadFileWorker.RunWorkerCompleted += LoadFileWorker_RunWorkerCompleted;

            saveFileWorker.WorkerReportsProgress = true;

            saveFileWorker.DoWork += SaveFileWorker_DoWork;
            saveFileWorker.ProgressChanged += SaveFileWorker_ProgressChanged;
            saveFileWorker.RunWorkerCompleted += SaveFileWorker_RunWorkerCompleted;
        }

        private void LoadFileWorker_DoWork(object? sender, DoWorkEventArgs e)
        {
            if (e.Argument is string arg)
            {
                ExcelReader er = new(arg);
                teachPlan = er.Plan;
                disciplines = er.Disciplines;
                compDisc = er.CompDisc;
            }
            else
            {
                throw new ArgumentException("e.Argument is not string");
            }
        }

        private void EnableControls()
        {
            loadFileButton.Enabled = true;
            saveFileButton.Enabled = true;
            disciplineCheckedList.Enabled = true;
        }

        private void DisableControls()
        {
            loadFileButton.Enabled = false;
            saveFileButton.Enabled = false;
            disciplineCheckedList.Enabled = false;
        }

        private void LoadFileWorker_RunWorkerCompleted(object? sender, RunWorkerCompletedEventArgs e)
        {
            void Clear()
            {
                teachPlan = null;
                disciplines = null;
                compDisc = null;
            }

            if (e.Error is ParseErrorException ex)
            {
                MessageBox.Show($"Произошла ошибка при разборе файла: {ex.Message}");

                Clear();
            }
            else if (e.Error is not null)
            {
                MessageBox.Show($"Internal error: {e.Error.Message}");

                Clear();
            }
            else
            {
                disciplineCheckedList.Items.Clear();
                foreach (var d in disciplines!)
                    disciplineCheckedList.Items.Add(d);
            }

            EnableControls();
        }

        private void SaveFileWorker_RunWorkerCompleted(object? sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Сгенерировано.");
            EnableControls();
            progressBar.Value = 0;
        }

        private void SaveFileWorker_DoWork(object? sender, DoWorkEventArgs e)
        {
            if (e.Argument is PlaceHolder ph && ph != null)
            {
                foreach (Discipline d in disciplineCheckedList.CheckedItems)
                {
                    ph.OneDiscipline(d);
                    saveFileWorker.ReportProgress(100 / disciplineCheckedList.CheckedItems.Count);
                }
            }
        }

        private void SaveFileWorker_ProgressChanged(object? sender, ProgressChangedEventArgs e)
        {
            progressBar.Step = e.ProgressPercentage;
            progressBar.PerformStep();
        }

        private void loadFileButton_Click(object sender, EventArgs e)
        {
            using OpenFileDialog ofd = new();

            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.DefaultExt = "xlsx";
            ofd.Filter = "Excel files (*.xlsx)|*.xlsx|All files|*.*";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                DisableControls();

                loadFileWorker.RunWorkerAsync(ofd.FileName);
            }
        }

        private void saveFileButton_Click(object sender, EventArgs e)
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
                DisableControls();

                PlaceHolder ph = new(teachPlan, compDisc, fbd.SelectedPath);
                saveFileWorker.RunWorkerAsync(ph);
            }
        }
    }
}