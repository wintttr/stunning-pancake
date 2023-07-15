using System.Windows.Forms;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using OpenFileDialog ofd = new();

            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.DefaultExt = "xlsx";
            ofd.Filter = "Excel files (*.xlsx)|*.xlsx|All files(*.*)|*.* ";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ExcelReader er = new(ofd.FileName);//@"C:\Users\wintttr\Downloads\Test.xlsx");
                var plan = er.GetPlan();
                teachPlan = plan;//вапроси?
                var disciplines = new List<string>(plan.Disciplines.Select(d => d.Name));

                checkedListBox1.Items.Clear();
                foreach (var d in disciplines)
                    checkedListBox1.Items.Add(d);
            }
        }

        public TeachPlan teachPlan;

        private void button2_Click(object sender, EventArgs e)
        {
            using FolderBrowserDialog fbd = new();

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                PlaceHolder ph = new PlaceHolder(teachPlan, fbd.SelectedPath);
            }
        }
    }
}