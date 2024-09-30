using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ReNameLestate
{
    public partial class Form1 : Form
    {

        private Button selectFileButton;
        private TextBox filePathTextBox;
        public string selectedFilePath;
        private Button selectFolderButton;
        private TextBox folderPathTextBox;
        public string selectedFolderPath;

        public Form1()
        {
            InitializeComponent();

            selectFileButton = new Button { Text = "Выбрать файл", Left = 10, Top = 20, Width = 100 };
            selectFileButton.Click += SelectFileButton_Click;
            filePathTextBox = new TextBox { Left = 120, Top = 20, Width = 500 };

            selectFolderButton = new Button { Text = "Выбрать папку", Left = 10, Top = 60, Width = 100 };
            selectFolderButton.Click += SelectFolderButton_Click;
            folderPathTextBox = new TextBox { Left = 120, Top = 60, Width = 500 };

            Controls.Add(selectFileButton);
            Controls.Add(filePathTextBox);

            Controls.Add(selectFolderButton);
            Controls.Add(folderPathTextBox);
        }

        private void SelectFileButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Все файлы (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = openFileDialog.FileName;
                    filePathTextBox.Text = selectedFilePath;
                }
            }
        }

        private void SelectFolderButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFolderPath = folderBrowserDialog.SelectedPath;
                    folderPathTextBox.Text = selectedFolderPath;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            string folderPath = selectedFolderPath;
            string excelFilePath = selectedFilePath;

            FileInfo fileInfo = new FileInfo(excelFilePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    textBox1.Text = "Лист Excel не найден!";
                    return;
                }

                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++)
                {
                    string oldFileBaseName = worksheet.Cells[row, 1].Text;
                    string newFileBaseName = worksheet.Cells[row, 2].Text;

                    string[] matchingFiles = Directory.GetFiles(folderPath, $"{oldFileBaseName}_*.jpg");

                    foreach (string oldFilePath in matchingFiles)
                    {
                        string fileName = Path.GetFileName(oldFilePath);

                        string suffix = fileName.Substring(oldFileBaseName.Length);

                        string newFilePath = Path.Combine(folderPath, newFileBaseName + suffix);

                        try
                        {
                            File.Move(oldFilePath, newFilePath);
                            textBox1.Text += $"Файл {fileName} переименован в {newFileBaseName}{suffix}{Environment.NewLine}";
                            textBox1.SelectionStart = textBox1.Text.Length;
                            textBox1.ScrollToCaret();
                        }
                        catch (Exception ex)
                        {
                            textBox1.Text += $"Ошибка при переименовании файла {fileName}: {ex.Message}";
                        }
                    }
                }
            }
            textBox1.Text += "Переименование файлов завершено.";
        }
    }
}
