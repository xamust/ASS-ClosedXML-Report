<b>2016 год.</b>

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

Выводы:
Собственно, как указывалось выше данный код полностью работоспособен, но написан быстро, для решения конкретной задачи на предприятии.
По хорошему, в данной задаче можно и нужно применить многопоточность. «Прикрутить» progressBar, сделать форму настроек для повышения юзабилити.
