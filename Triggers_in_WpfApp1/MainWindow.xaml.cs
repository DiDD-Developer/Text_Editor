using iTextSharp.text.pdf;
using iTextSharp.text;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Triggers_in_WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string fileName = openFileDialog.FileName;
                    using (StreamReader reader = new StreamReader(fileName))
                    {
                        if (!string.IsNullOrEmpty(TextBox1.Text))
                        {

                        }
                        else
                        {
                            TextBox1.Text = reader.ReadToEnd();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при открытии файла: " + ex.Message);
                }
            }
        }

        private void UpdateTextBoxFontStyle()
        {
            FontStyle style = FontStyles.Normal;
            FontWeight weight = FontWeights.Normal;
            TextDecorationCollection decorations = new TextDecorationCollection();

            if (BoldCheckBox.IsChecked == true)
            {
                weight = FontWeights.Bold;
            }

            if (ItalicCheckBox.IsChecked == true)
            {
                style = FontStyles.Italic;
            }

            if (DeleteCheckBox.IsChecked == true)
            {
                decorations.Add(TextDecorations.Underline);
            }

            TextBox1.FontStyle = style;
            TextBox1.FontWeight = weight;
            TextBox1.TextDecorations = decorations;
        }

        private void BoldCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            UpdateTextBoxFontStyle();
        }

        private void BoldCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateTextBoxFontStyle();
        }

        private void ItalicCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            UpdateTextBoxFontStyle();
        }

        private void ItalicCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateTextBoxFontStyle();
        }

        private void DeleteCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            UpdateTextBoxFontStyle();
        }

        private void DeleteCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateTextBoxFontStyle();
        }

        private void FontSizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Обновление размера шрифта в текстовом поле
            ComboBox comboBox = (ComboBox)sender;
            if (comboBox.SelectedItem != null)
            {
                double fontSize = double.Parse(((ComboBoxItem)comboBox.SelectedItem).Content.ToString());
                TextBox1.FontSize = fontSize;
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            TextBox1.Clear();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TextBox1.Text))
            {
                Close();
            }
            else
            {
                MessageBox.Show("Чтобы закрыть программу, вначале очистите поле для текста кнопкой 'Очистить' сверху!", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

        }

        private void SaveAsPDF(string text, string fileName, string fontFamily, double fontSize, bool isItalic, bool isBold)
        {
            MessageBox.Show($"fontFamily: {fontFamily}, fontSize: {fontSize}, isItalic: {isItalic}, isBold: {isBold}");

            using (FileStream fs = new FileStream(fileName, FileMode.Create))
            {
                iTextSharp.text.Document document = new iTextSharp.text.Document();
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                document.Open();

                iTextSharp.text.Font font = iTextSharp.text.FontFactory.GetFont(fontFamily, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                font.Size = (float)fontSize;

                if (isItalic)
                {
                    font.SetStyle("Italic");
                    Console.WriteLine("Font style: Italic");
                }
                if (isBold)
                {
                    font.SetStyle("Bold");
                    Console.WriteLine("Font style: Bold");
                }

                iTextSharp.text.Paragraph paragraph = new iTextSharp.text.Paragraph(text, font);

                document.Add(paragraph);
                document.Close();
            }
        }

        private void SaveAsPDF_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == true)
            {
                string fileName = saveFileDialog.FileName;
                string fontFamily = string.Empty;

                double fontSize = 12;
                fontSize = double.Parse(((ComboBoxItem)FontSizeComboBox.SelectedItem).Content.ToString());
                bool isItalic = ItalicCheckBox.IsChecked == true;
                bool isBold = BoldCheckBox.IsChecked == true;

                if (DeleteCheckBox.IsChecked == true)
                {
                    fontFamily = "Arial";
                }
                else if (ItalicCheckBox.IsChecked == true && BoldCheckBox.IsChecked == true)
                {
                    fontFamily = "Arial Bold Italic";
                }
                else if (ItalicCheckBox.IsChecked == true)
                {
                    fontFamily = "Arial Italic";
                }
                else if (BoldCheckBox.IsChecked == true)
                {
                    fontFamily = "Arial Bold";
                }
                else
                {
                    fontFamily = "Arial";
                }

                SaveAsPDF(TextBox1.Text, fileName, fontFamily, fontSize, isItalic, isBold);
                MessageBox.Show("Файл успешно сохранен как PDF.");
            }
        }

        private void SaveAsTxt(string text, string fileName)
        {
            File.WriteAllText(fileName, text);
        }

        private void SaveAsTxt_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == true)
            {
                string fileName = saveFileDialog.FileName;
                SaveAsTxt(TextBox1.Text, fileName);
                MessageBox.Show("Файл успешно сохранен как TXT.");
            }
        }

        private void SaveAsDocx(string text, string fileName, string fontFamily, double fontSize, bool isItalic, bool isBold, bool isUnderline)
        {
            using (DocumentFormat.OpenXml.Packaging.WordprocessingDocument wordDocument = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(fileName, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                DocumentFormat.OpenXml.Wordprocessing.Document doc = new DocumentFormat.OpenXml.Wordprocessing.Document();
                DocumentFormat.OpenXml.Wordprocessing.Body body = new DocumentFormat.OpenXml.Wordprocessing.Body();
                DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();

                DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run();

                DocumentFormat.OpenXml.Wordprocessing.RunProperties runProps = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();

                if (isBold)
                {
                    runProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Bold());
                }

                if (isItalic)
                {
                    runProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Italic());
                }

                if (isUnderline)
                {
                    runProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Underline() { Val = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single });
                }

                runProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.RunFonts() { Ascii = fontFamily });
                runProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = $"{fontSize * 2}" });

                run.AppendChild(runProps);
                run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(text));

                paragraph.AppendChild(run);
                body.AppendChild(paragraph);
                doc.AppendChild(body);
                mainPart.Document = doc;
                mainPart.Document.Save();
            }
        }

        private void SaveAsDocx_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word files (*.docx)|*.docx|All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == true)
            {
                string fileName = saveFileDialog.FileName;
                string fontFamily = string.Empty;

                double fontSize = 12;
                fontSize = double.Parse(((ComboBoxItem)FontSizeComboBox.SelectedItem).Content.ToString());
                bool isItalic = ItalicCheckBox.IsChecked == true;
                bool isBold = BoldCheckBox.IsChecked == true;
                bool isUnderline = DeleteCheckBox.IsChecked == true; // Проверяем, выбран ли чекбокс подчеркивания

                if (DeleteCheckBox.IsChecked == true)
                {
                    fontFamily = "Arial";
                }
                else if (ItalicCheckBox.IsChecked == true && BoldCheckBox.IsChecked == true)
                {
                    fontFamily = "Arial Bold Italic";
                }
                else if (ItalicCheckBox.IsChecked == true)
                {
                    fontFamily = "Arial Italic";
                }
                else if (BoldCheckBox.IsChecked == true)
                {
                    fontFamily = "Arial Bold";
                }
                else
                {
                    fontFamily = "Arial";
                }

                SaveAsDocx(TextBox1.Text, fileName, fontFamily, fontSize, isItalic, isBold, isUnderline); // Передаем isUnderline в качестве аргумента
                MessageBox.Show("Файл успешно сохранен как DOCX.");
            }
        }
    }
}
