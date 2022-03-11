ClosedXML Excel C#
Лирическое отступление:
"Поскольку профессия администратора сетей многогранная и творческая...."

Вводная информация:
Имеем 2 входных файла Excel, пусть ФАЙЛ_1 и ФАЙЛ_2, каждый имеет по 10-15 столбцов и порядка 10 – 15 тысячи строк информации.

Постановка задачи:
Необходимо сравнивать в ФАЙЛе_2, НЕ желтые строки по столбцу “С” со строкой по комбинации столбцов “B”+”C"+”D” ФАЙЛа_1. В случае совпадения строки в ФАЙЛе_1, шрифт меняем на красный.

Краткие размышления:
Необходимо обработать 10 тыс. * 10 тыс. = 100 млн. комбинаций (приблизительно).
Использовать средства разработки MSOffice для решения этой задачи нерационально.
Логично и правильно использовать OpenXML, поэтому используем упрощенную версию, ClozedXML. 
Качаем отсюда, там же берем документацию и ищем ответы на возникающие вопросы. (upd: исходники переехали на https://github.com/ClosedXML/ClosedXML)


Исполнение:
//Открываем наши Excel файлы
excelApp1XML = new ClosedXML.Excel.XLWorkbook(file1);
excelApp2XML = new ClosedXML.Excel.XLWorkbook(file2);

//Выбираем лист с которым будем работать.
//Делал быстро «на коленке», поэтому и (1)
sheetExcel1XML = excelApp1XML.Worksheet(1);
sheetExcel2XML = excelApp2XML.Worksheet(1);

//Считаем количество используемых строк на листе
XMLlastRow1 = sheetExcel1XML.RowsUsed().Count();
XMLlastRow2 = sheetExcel2XML.RowsUsed().Count();

//Для удобства сравнения двух выборок используем System.Collections.Generic.Dictionary<TKey, TValue>
//В качестве ключа примем номер строки
 var myDictionary1 = new Dictionary<int, string>();
 var myDictionary2 = new Dictionary<int, string>();

//В myDictionary2 записываем значения из столбца “C” ФАЙЛа_2
//В myDictionary1 записываем комбинацию столбцов “B”+”C"+”D” ФАЙЛа_1
             
 for (int r = 1; r <= XMLlastRow2; r++)
                {
                    if (sheetExcel2XML.Cell(r, "C").Style.Fill.BackgroundColor.ToString() == "FFFFFFFF")
                    {
                        string tr = sheetExcel2XML.Cell(r, "C").Value.ToString();
                        string tr2;

                        // для добавления  бланков с первыми 000

                        char[] arry = tr.ToCharArray();
                        if (arry.Length < 11)
                        {
                            int sdvig = 11 - arry.Length;
                            char[] arry_new = new char[sdvig];
                            for (int f = 0; f < sdvig; f++)
                            {
                                arry_new[f] = '0';
                            }
                            tr2 = new string(arry_new);
                            myDictionary2.Add(r, tr2 + tr);
                        }
                        else
                        {
                            myDictionary2.Add(r, sheetExcel2XML.Cell(r, "C").Value.ToString());
                        }
                    }
                    else
                    {
                        myDictionary2.Add(r, "000");
                    }
                }
                for (int i = 1; i <= XMLlastRow1; i++)
                {
                    string final = "";
                    string str = sheetExcel1XML.Cell(i, "C").Value.ToString();
                    string str2 = sheetExcel1XML.Cell(i, "B").Value.ToString();
                    string str3 = sheetExcel1XML.Cell(i, "D").Value.ToString();
                    char[] arr = str.ToCharArray();
                    char[] arr_new = new char[7];

                    for (int j = 1; j <= 7; j++)
                    {
                        try
                        {
                            arr_new[j - 1] = arr[j];
                        }
                        catch (Exception ecr)
                        {

                        }
                    }
                    final = new string(arr_new);
                    myDictionary1.Add(i, str2 + final + str3);
                }
//Поскольку ФАЙЛ_2 больше нам не нужен, то освободим ресурсы занятые им и спокойно про него забудем
sheetExcel2XML.Dispose();
excelApp2XML.Dispose();

//Изменения в ФАЙЛе_2 будут происходить постоянно (желтые строки могут стать белыми и наоборот, как решит вселенная), поэтому надо создать некоторого рода проверку на изменения.
//Пойдем по простому пути, при запуске программы, все значения в ФАЙЛе_1 будут окрашиваться по умолчанию в черный цвет
sheetExcel1XML.Style.Font.FontColor = ClosedXML.Excel.XLColor.Black;

//Сравним значения двух Dictionary (ФАЙЛ_1 и ФАЙЛ_2), в случае совпадения  красим шрифт красным в ФАЙЛе_1
       int ground = 0;
                for (int i = 1; i <= XMLlastRow1; i++)
                {
                    string out1 = "";
                    myDictionary1.TryGetValue(i, out out1);
                    for (int co = 1; co < XMLlastRow2; co++)
                    {
                        string out2 = "";
                        myDictionary2.TryGetValue(co, out out2);
                        if (out1.Equals(out2))
                        {
                            sheetExcel1XML.Row(i).Style.Font.FontColor = ClosedXML.Excel.XLColor.Red;
                            ground++;
                        }
                    }
                }

//Сохраняем наш ФАЙЛ_1
excelApp1XML.Save();

Выводы:
Собственно, как указывалось выше данный код полностью работоспособен, но написан быстро, для решения конкретной задачи на предприятии.
По хорошему, в данной задаче можно и нужно применить многопоточность. «Прикрутить» progressBar, сделать форму настроек для повышения юзабилити.
