using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;
using Word1 = Microsoft.Office.Tools.Word;
using System.Reflection;



namespace DataGrid
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Word._Application application;
        Word._Document document;
        public MainWindow()
        {
            InitializeComponent();          
        }

        //Загрузка содержимого таблицы
        private void grid_Loaded(object sender, RoutedEventArgs e)
        {
            
            object oMissing = Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
            //Start Word and create a new document.
            Word._Application oWord = new Word.Application();
            oWord.Visible = true;
            Word._Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            //Word.Table f = new Table(@"C:\Users\Йонас\Desktop\1.docx");
            //MessageBox.Show(f.Tables);

            

            // loop through the rows in the table and send the contents of the row
            //  to the DataGridView
            List<MyTable> result = new List<MyTable>(2);
            result.Add(new MyTable("Майкл Джексон", "Thriller"));
            result.Add(new MyTable("AC/DC", "Back in Black"));
            result.Add(new MyTable("AC/DC", "Back in Black"));
            result.Add(new MyTable("AC/DC", "Back in Black"));
            result.Add(new MyTable("AC/DC", "Back in Black"));
            result.Add(new MyTable("AC/DC", "Back in Black"));

            
            Object missingObj = Missing.Value;
            Object trueObj = true;
            Object falseObj = false;

            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            Object templatePathObj = @"C:\Users\Йонас\Desktop\матпомощь\1.docx";

            // если вылетим не этом этапе, приложение останется открытым
            try
            {
                document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (Exception error)
            {
                document.Close(ref falseObj, ref missingObj, ref missingObj);
                application.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                application = null;
                throw error;
            }
            application.Visible = true;

            //firstTable = new Word1.Document();
            // foreach (Word1.Row row in firstTable.Rows)
            // {
            //     List<string> cellValues = new List<string>();
            //     foreach (Word1.Cell cell in row.Cells)
            //     {
            //         string cellContents = cell.Range.Text;
            //         MessageBox.Show(cellContents);
            //     }
            // }



            grid.ItemsSource = result;


            MessageBox.Show("Действие выполнено");
        }

        //Получаем данные из таблицы
        private void grid_MouseUp(object sender, MouseButtonEventArgs e)
        {
        }
    }


    class MyTable
    {
        public MyTable(string defenition, string examplecode)
        {
            this.Defenition = defenition;
            this.Examplecode = examplecode;
        }
        public string Examplecode { get; set; }
        public string Defenition { get; set; }
    }

}
