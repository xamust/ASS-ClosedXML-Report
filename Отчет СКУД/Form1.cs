using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using ClosedXML.Excel;
using System.Threading;

namespace Отчет_СКУД
{
    public partial class Form1 : Form
    {
        public string file1 = null;
        public string file2 = null;
        public string file3 = null;
        public string file4 = null;

        int XMLlastRow1 = 0;
        int XMLlastRow2 = 0;
        int XMLlastRow3 = 0;
        int XMLlastRow4 = 0;

        ClosedXML.Excel.XLWorkbook excelApp1XML = null;
        ClosedXML.Excel.XLWorkbook excelApp2XML = null;
        ClosedXML.Excel.XLWorkbook excelApp3XML = null;
        ClosedXML.Excel.XLWorkbook excelApp4XML = null;

        ClosedXML.Excel.IXLWorksheet sheetExcel1XML = null;
        ClosedXML.Excel.IXLWorksheet sheetExcel2XML = null;
        ClosedXML.Excel.IXLWorksheet sheetExcel3XML = null;
        ClosedXML.Excel.IXLWorksheet sheetExcel4XML = null;
       
        public Form1()
        {
            InitializeComponent();
            this.Text = "Отчет СКУД v1.0.1";
            this.label1.Text = "Форма отчета";
            this.label2.Text = "Файл СКУД";
            this.label3.Text = "1C Кадры";
            this.label4.Text = "ТУРВ";
            this.label5.Text = "Готов к работе";
            this.button1.Text = "Загрузка";
            this.button2.Text = "Загрузка";
            this.button3.Text = "Загрузка";
            this.button4.Text = "Загрузка";
            this.textBox1.Enabled = false;
            this.textBox2.Enabled = false;
            this.textBox3.Enabled = false;
            this.textBox4.Enabled = false;
            progressBar1.Maximum = 6;
            progressBar1.Minimum = 1;
            progressBar1.Step = 1;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Книга Excel |*.xlsx|Книга Microsoft Excel 97 — 2003 |*.xls|Книга Excel 4.0 |*.xlw|Лист Excel (код) |*.xlsm";
            openFileDialog1.Title = "Select a Excel File";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                file1 = textBox1.Text;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Книга Excel |*.xlsx|Книга Microsoft Excel 97 — 2003 |*.xls|Книга Excel 4.0 |*.xlw|Лист Excel (код) |*.xlsm";
            openFileDialog2.Title = "Select a Excel File";
            if (openFileDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = openFileDialog2.FileName;
                file2 = textBox2.Text;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog3 = new OpenFileDialog();
            openFileDialog3.Filter = "Книга Excel |*.xlsx|Книга Microsoft Excel 97 — 2003 |*.xls|Книга Excel 4.0 |*.xlw|Лист Excel (код) |*.xlsm";
            openFileDialog3.Title = "Select a Excel File";
            if (openFileDialog3.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox3.Text = openFileDialog3.FileName;
                file3 = textBox3.Text;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog4 = new OpenFileDialog();
            openFileDialog4.Filter = "Книга Excel |*.xlsx|Книга Microsoft Excel 97 — 2003 |*.xls|Книга Excel 4.0 |*.xlw|Лист Excel (код) |*.xlsm";
            openFileDialog4.Title = "Select a Excel File";
            if (openFileDialog4.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox4.Text = openFileDialog4.FileName;
                file4 = textBox4.Text;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (file1 != null || file2 != null || file3 != null || file4 != null)
            {
                progressBar1.Value = 1;
                Thread myThread = new Thread(myThreader);
                myThread.IsBackground = true;
                myThread.Start();
            }
            else MessageBox.Show("Введены некорректные данные!");
        }

        void myThreader() //поток
        {
            this.Invoke(new ThreadStart(delegate { 
                progressBar1.PerformStep();
                this.label5.Text = "Обработка файла СКУД...";
                this.button1.Enabled = false;
                this.button2.Enabled = false;
                this.button3.Enabled = false;
                this.button4.Enabled = false;
                this.button9.Enabled = false;
            }));
            Module1(file1, file2);
            this.Invoke(new ThreadStart(delegate { 
                progressBar1.PerformStep();
                this.label5.Text = "Обработка файла 1С Кадры...";
            }));
            Module2(file1, file3);
            this.Invoke(new ThreadStart(delegate { 
                progressBar1.PerformStep();
                this.label5.Text = "Обработка файла ТУРВ...";
            }));
            Module3(file1, file4);
            this.Invoke(new ThreadStart(delegate { 
                progressBar1.PerformStep();
                this.label5.Text = "Форматирование файла отчета...";
            }));
            Module4(file1);
            this.Invoke(new ThreadStart(delegate
            {
                progressBar1.PerformStep();
                this.label5.Text = "Готово";
                this.button1.Enabled = true;
                this.button2.Enabled = true;
                this.button3.Enabled = true;
                this.button4.Enabled = true;
                this.button9.Enabled = true;
            }));
        }

        String NameConvertion(String name)
        {
            String[] mass1 = name.Split(' ');
            String result = ""; 

            int count = 0;
            for (int i = 0; i < mass1.Length; i++)
            {
                //Отработка пробелов фамилии
                if (mass1[i].Length != 0 & count < 1)
                {
                    result = mass1[i] + " ";
                    count++;
                }
                //Отработка пробелов инициалов
                else if (mass1[i].Length != 0 & count >= 1 & count<= 2)
                {
                   result = result + mass1[i].Substring(0, 1) + ".";
                   count++;
                }
            }
            return result;           
        }

        String NameCanonical(String name)
        {
            String resultC = "";
            String[] mass2 = name.Split(' ');
            int count = 0;

            if (!name.Contains("."))
            {
                resultC = name; //корейские  фамилии (и пр. без точки)
            }
            else
            {

                for (int i = 0; i < mass2.Length; i++)
                {
                    if (mass2[i].Length != 0 & count < 1)
                    {
                        resultC = mass2[i] + " ";
                        count++;
                    }

                    else if (mass2[i].Length != 0 & count >= 1)
                    {
                        resultC = resultC + mass2[i].Substring(0, 4);
                        count++;
                    }
                }
            }
                return resultC;
        }

        void CreateDoc()
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Сводка за __-__.__. " + new DateTime().Year.ToString());
            worksheet.Cell("A1").Value = "Hello World!";
            workbook.SaveAs("Temp.xlsx");
        }

        //Module1 - done
        void Module1(String file1, String file2)
        {
            
            try {
              //  CreateDoc();
                excelApp1XML = new ClosedXML.Excel.XLWorkbook(file1);
                excelApp2XML = new ClosedXML.Excel.XLWorkbook(file2);
                sheetExcel1XML = excelApp1XML.Worksheet(1);
                sheetExcel2XML = excelApp2XML.Worksheet(1);
                XMLlastRow2 = sheetExcel2XML.RowsUsed().Count();

                //Переносим примечание
                var firstTableCell1 = sheetExcel1XML.Cell(7, "G");
                var lastTableCell1 = sheetExcel1XML.Cell(10, "I");
                var rngData1 = sheetExcel1XML.Range(firstTableCell1.Address, lastTableCell1.Address);
                sheetExcel1XML.Cell(XMLlastRow2+3, "G").Value = rngData1; //отступ в 3 строки
                rngData1.Value = ""; //затираем изначальные данные
                
                
                //Переносим данные 1

                // Определяем область для копирования
              //  var firstTableCell = sheetExcel2XML.FirstCellUsed(); //индекс первой заполненной ячейки
                var firstTableCell = sheetExcel2XML.Cell(5,"A");
                // var lastTableCell = sheetExcel2XML.LastCellUsed(); //индекс последней заполненной ячейки
                var lastTableCell = sheetExcel2XML.Cell(XMLlastRow2+4, "C");
                var rngData = sheetExcel2XML.Range(firstTableCell.Address, lastTableCell.Address);

                // Copy the table to another worksheet
             //   var wsCopy = workbook.Worksheets.Add("Contacts Copy");
                sheetExcel1XML.Cell(2, "B").Value = rngData;
                
                //Переносим данные 2
                var firstTableCell2 = sheetExcel2XML.Cell(5, "F");
                var lastTableCell2 = sheetExcel2XML.Cell(XMLlastRow2 + 4, "F");
                var rngData2 = sheetExcel2XML.Range(firstTableCell2.Address, lastTableCell2.Address);
                sheetExcel1XML.Cell(2, "H").Value = rngData2;

                //Переносим данные 3
                var firstTableCell3 = sheetExcel2XML.Cell(5, "E");
                var lastTableCell3 = sheetExcel2XML.Cell(XMLlastRow2 + 4, "E");
                var rngData3 = sheetExcel2XML.Range(firstTableCell3.Address, lastTableCell3.Address);
                sheetExcel1XML.Cell(2, "I").Value = rngData3;

                //Переносим данные 4
                var firstTableCell4 = sheetExcel2XML.Cell(5, "G");
                var lastTableCell4 = sheetExcel2XML.Cell(XMLlastRow2 + 4, "G");
                var rngData4 = sheetExcel2XML.Range(firstTableCell4.Address, lastTableCell4.Address);
                sheetExcel1XML.Cell(2, "J").Value = rngData4;

              //сохраняем
                excelApp1XML.Save();
               // MessageBox.Show(XMLlastRow2.ToString());
                sheetExcel2XML.Dispose();
                excelApp2XML.Dispose();
                File.Delete(file2);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());    
            }
        }
        //Module2 - done
        void Module2(String file1, String file3)
        {
            String memoryOrg = "";
            try {
                excelApp1XML = new ClosedXML.Excel.XLWorkbook(file1);
                excelApp3XML = new ClosedXML.Excel.XLWorkbook(file3);
                sheetExcel1XML = excelApp1XML.Worksheet(1);
                sheetExcel3XML = excelApp3XML.Worksheet(1);
                XMLlastRow1 = sheetExcel1XML.RowsUsed().Count();
                XMLlastRow3 = sheetExcel3XML.RowsUsed().Count();

                for (int i = 1; i <= XMLlastRow1; i++)
                {
                    if (sheetExcel1XML.Cell(i, "B").Value.ToString() != "Сотрудник")
                    {
                        for (int r = 1; r <= XMLlastRow3; r++)
                        {
                            //вытаскиваем название службы (если ячейка I не объед. и значение A не 0 и J не объед., то нам подходит)
                            if (sheetExcel3XML.Cell(r, "I").IsMerged() & sheetExcel3XML.Cell(r, "A").Value.ToString() != "" & !sheetExcel3XML.Cell(r, "J").IsMerged())
                            {
                                // MessageBox.Show("Служба " + sheetExcel3XML.Cell(r, "A").Value.ToString()); //for test
                             //   memoryOrg = sheetExcel3XML.Cell(r, "A").Value.ToString();
                            }
                            //для отлавливания названия отдела
                            else if (sheetExcel3XML.Cell(r, "I").IsMerged() & sheetExcel3XML.Cell(r, "A").Value.ToString() == "")
                            {
                                // MessageBox.Show("Отдел " + sheetExcel3XML.Cell(r, "B").Value.ToString()); //for test
                                memoryOrg = sheetExcel3XML.Cell(r, "B").Value.ToString();
                            }
                            //вытаскиваем ФИО
                            else if (!sheetExcel3XML.Cell(r, "F").IsMerged() & sheetExcel3XML.Cell(r, "F").Value.ToString() != "Фамилия Имя Отчество")
                            {
                                if (NameCanonical(sheetExcel1XML.Cell(i, "B").Value.ToString()).Equals(NameConvertion(sheetExcel3XML.Cell(r, "F").Value.ToString())))
                                {
                                   // this.Invoke(new ThreadStart(delegate { label6.Text = sheetExcel1XML.Cell(i, "B").Value.ToString() + " " + sheetExcel3XML.Cell(r, "F").Value.ToString(); }));
                                    sheetExcel1XML.Cell(i, "K").Value = sheetExcel3XML.Cell(r, "G").Value.ToString();//Должность
                                    sheetExcel1XML.Cell(i, "L").Value = memoryOrg;//Подразделение
                                }
                            }
                        }
                    }
                }
                //сохраняем
                excelApp1XML.Save();
                //удаляем
                sheetExcel3XML.Dispose();
                excelApp3XML.Dispose();
                File.Delete(file3);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        //Module3 - done
        void Module3(String file1, String file4)
        {
            try
            {
                excelApp1XML = new ClosedXML.Excel.XLWorkbook(file1);
                excelApp4XML = new ClosedXML.Excel.XLWorkbook(file4);
                sheetExcel1XML = excelApp1XML.Worksheet(1);
                sheetExcel4XML = excelApp4XML.Worksheet(1);
                XMLlastRow1 = sheetExcel1XML.RowsUsed().Count();
                XMLlastRow4 = sheetExcel4XML.RowsUsed().Count();
               
                //цикл выставления Дня недели
                for (int i = 2; i <= XMLlastRow1; i++)
                {
                   String[] mass21 = sheetExcel1XML.Cell(i,"C").Value.ToString().Split('.');
                   if (mass21.Length != 3) break;
                   DateTime dt = new DateTime(Convert.ToInt32(mass21[2]),Convert.ToInt32(mass21[1]),Convert.ToInt32(mass21[0]));
                  //for test, для CultureInfo необходимо подключить System.Globalization
                   sheetExcel1XML.Cell(i, "E").Value = dt.ToString("dddd", CultureInfo.GetCultureInfo("ru-ru"));
                }  
                
                for (int i = 1; i <= XMLlastRow1; i++)
                {
                    for (int r = 2; r <= XMLlastRow4; r++)
                    {
                        if(NameCanonical(sheetExcel1XML.Cell(i, "B").Value.ToString()).Equals(NameConvertion(sheetExcel4XML.Cell(r, "B").Value.ToString())))
                        {
                            
                            String[] massint = sheetExcel1XML.Cell(i, "C").Value.ToString().Split('.');
                             for (int j = 4; j <= 34; j++)
                            {
                                if (NameCanonical(sheetExcel1XML.Cell(i, "B").Value.ToString()).Equals("Чеботрева Т.Г.")) MessageBox.Show(sheetExcel4XML.Cell(r, j).Value.ToString());
                                 //for testing
                               /* MessageBox.Show("Дата из отчета: " +
                                                sheetExcel1XML.Cell(i, "C").Value.ToString() +
                                                '\n' +
                                                "Дата из отчета: " +
                                                massint[0] + 
                                                '\n' +
                                                "День из ТУРВа: " +
                                                sheetExcel4XML.Cell(1, j).Value.ToString() + 
                                                '\n' +
                                                "Совпадение: " +
                                                Convert.ToInt32(massint[0]).ToString().Equals(sheetExcel4XML.Cell(1, j).Value.ToString())
                                                );
                                 */
                                if (Convert.ToInt32(massint[0]).ToString().Equals(sheetExcel4XML.Cell(1, j).Value.ToString()))
                                {
                                    sheetExcel1XML.Cell(i, "F").Value = sheetExcel4XML.Cell(r, j).Value.ToString();
                                    /* String result = sheetExcel4XML.Cell(r, j).Value.ToString().Split('\n')[0]; // отделяем новую строку и берем первый индекс  
                                      if (result.Length >= 3 && result.Substring(0, 1).Equals("Я"))  //Ищем Я (явки) убираем Я, оставляем числовое значение, отправляем в файл отчета
                                      {
                                          // MessageBox.Show("Length of result: " + result.Length.ToString() + ". Value of result: " + result + ". Substring(2,length-2) result: " + result.Substring(2, result.Length - 2));  // проверка
                                          sheetExcel1XML.Cell(i, "F").Value = result.Substring(2, result.Length - 2);
                                      }
                                      else //Вставляем оставшиеся, только буквы
                                      {
                                          if (result.Length > 1) sheetExcel1XML.Cell(i, "F").Value = result.Substring(0, 2);
                                          else sheetExcel1XML.Cell(i, "F").Value = result;                                      
                                      }*/

                                    //проверяем на совместительство
                                    if (NameConvertion(sheetExcel4XML.Cell(r, "B").Value.ToString()).Equals(NameConvertion(sheetExcel4XML.Cell(r + 1, "B").Value.ToString())))
                                    {
                                        sheetExcel1XML.Cell(i, "G").Value = sheetExcel4XML.Cell(r + 1, j).Value.ToString();
                                        /* String result2 = sheetExcel4XML.Cell(r+1, j).Value.ToString().Split('\n')[0]; // отделяем новую строку и берем первый индекс
                                                                          if (result2.Length >= 3 && result2.Substring(0, 1).Equals("Я"))  //Ищем Я (явки) убираем Я, оставляем числовое значение, отправляем в файл отчета
                                          {
                                              sheetExcel1XML.Cell(i, "G").Value = result2.Substring(2, result2.Length - 2);
                                          }
                                          else //Вставляем оставшиеся, только буквы
                                          {
                                              if (result2.Length > 1) sheetExcel1XML.Cell(i, "G").Value = result2.Substring(0, 2);
                                              else sheetExcel1XML.Cell(i, "G").Value = result2;
                                          }*/
                                        r++;
                                    }
                                }

                            }
                        }
                    }
                }

                //Форматируем ячейки
                var firstTableCell11 = sheetExcel1XML.Cell(2, "F");
                var lastTableCell11 = sheetExcel1XML.Cell(XMLlastRow1 - 4, "G");
                var rngData1 = sheetExcel1XML.Range(firstTableCell11.Address, lastTableCell11.Address);
                rngData1.Style.NumberFormat.NumberFormatId = 1; 

                //сохраняем
                excelApp1XML.Save();
                //удаляем
                sheetExcel4XML.Dispose();
                excelApp4XML.Dispose();
                File.Delete(file4);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        //Module4 - done
        void Module4(String file1)
        {
            try
            {
                excelApp1XML = new ClosedXML.Excel.XLWorkbook(file1);
                sheetExcel1XML = excelApp1XML.Worksheet(1);
                XMLlastRow1 = sheetExcel1XML.RowsUsed().Count();
                

                

                var firstTableCell11 = sheetExcel1XML.Cell(2, "A");
                var lastTableCell11 = sheetExcel1XML.Cell(XMLlastRow1, "L");
                var rngData1 = sheetExcel1XML.Range(firstTableCell11.Address, lastTableCell11.Address);

                //шрифт
                rngData1.Style.Font.FontColor = ClosedXML.Excel.XLColor.Black;
                rngData1.Style.Font.FontName = "Times New Roman";
                rngData1.Style.Font.FontSize = 10;
                //фон
                rngData1.Style.Fill.SetBackgroundColor(ClosedXML.Excel.XLColor.NoColor);
                //границы ячеек
                rngData1.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                rngData1.Style.Border.DiagonalBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                rngData1.Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                rngData1.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                rngData1.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                rngData1.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                rngData1.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                //выравнивание текста в ячейке
                rngData1.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                rngData1.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Top;
                //отступ
                rngData1.Style.Alignment.Indent = 0; 
                //перенос по словам
                rngData1.Style.Alignment.WrapText = false;
                //высота ячейки
                sheetExcel1XML.Rows(2, XMLlastRow1).Height = 12.5;
                //выравнивание столбца по ширине
                sheetExcel1XML.Columns().AdjustToContents();
                //проставляем  № п/п
                for (int i = 2; i <= XMLlastRow1 - 4; i++) { sheetExcel1XML.Cell(i, "A").Value = i - 1; } //-4, доп. ячейки с примечанием
                //сохраняем
                excelApp1XML.Save();
                this.Invoke(new ThreadStart(delegate
                {
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Книга Excel |*.xlsx";
                    saveFileDialog1.Title = "Select a Excel File";
                    if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK) excelApp1XML.SaveAs(saveFileDialog1.FileName);
                }));

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        } 
    }
}
