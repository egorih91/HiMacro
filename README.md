# HiMacro

Данный макрос разработан для приветствия адресата (адресатов) нового исходящего письма.

!!!!
Макрос не затрагивает системные файлы, не влияет на работоспособность операционной системы, имеет открытый код.
Данное утверждение верно если дата изменения файла - 12.10.2018.
_________________________________________________________________________________________________________________

Для установки макроса необходимо выполните следующие действия:
1. В Outlook включить вкладку "Разработчик". Для этого необходимо нажать "Файл"-"Параметры"-"Настроить ленту".
В окне основные вкладки нужно включить галочку напротив пункта "Разработчик". И применить настройки.
2. После чего надо перейти во вкладку "Разработчик".
3. Нажать "Безопасность макросов" - "Включить все макросы" - "ОК". 
После данного пункта во избежание проблем с работой макроса рекомендуется перезагрузка компьютера.
4. Нажать "Visual Basic". В окошке Project сделать клик ПКМ (правой клавишей мыши) и нажать "Import file".
5. В открывшемся окне выбрать файл из данного архива с расширением .bas и нажать "Открыть".
6. После чего нажать "File"-"Save VbaProject.OTM (или знак дискетки для сохранения).
7. Закрыть окно редактора Microsoft Visual Basic for Application.
8. Создать новое исходящее письмо клавишей "Создать сообщение".
9. В верхней строке после команд "Сохранить", "Отменить" и т.п. нажать "Настроить панель быстрого доступа".
10. В открывшемся списке выбрать "Другие команды".
11. В поле "Выбрать команды из" выбрать "Макросы".
12. Поочередно добавить "Проект1.FIOMAIN" и "Проект1.IOFMAIN", нажать "ОК".
(можно отредактировать значки и наименование быстрых команд выбрав их в окошке справа и нажав кнопку "Изменить")

Две новых команды появятся на панели быстрого доступа. 

Первая из них (Проект1.FIOMAIN) должна использоваться если адресат записан в формате Фамилия Имя Отчество(не обязательно).
Вторая из них (Проект1.IOFMAIN) должна использоваться если адресат записан в формате Имя Отчество(не обязательно) Фамилия.

----------------------------------------------------------------------------------------------------------------------------

БОЛЕЕ ПОДРОБНОЕ ОПИСАНИЕ РАБОТЫ МАКРОСА

Данный макрос по нажатию нужно быстрой команды (в зависимости от формата написания имени адресата)
вписывает в тело письма обращение и приветствие.  

Если адресатов несколько - приветствие имеет следующий формат "Уважаемые коллеги, добрый день!"

В случае наличия отчества адресата сообщение корректируется в зависимости от пола получателя и имеет вид
"Уважаемый Имя Отчество(адресата), добрый день!" или "Уважаемая Имя Отчество(адресата), добрый день!"

В случае отсутствия отчества обращение будет формата "Имя(адресата), добрый день!"

Само приветствие "добрый день!" можно заменить на любое другое изменив в коде значение переменной "fin".
Для этого после установки, описанной выше, нужно выполнить из неё шаги 2,4, после чего 
открыть "Проект 1"-"Modules"-имеющийся там файл (сделать двойной клик по нему). 
Найти строку fin = "добрый день!" и заменить то, что заключено в двойные кавычки на нужное значение.

----------------------------------------------------------------------------------------------------------------------------
Ограничения:
1.Макрос и данный мануал написан для Microsoft Outlook 2016 (64bit). 
Работоспособность макроса и пошаговое совпадение инструкции на других версиях программы не гарантируется.

В случае наличия замечаний, обнаружения некорректной работы или предложений обращаться на egorih91@gmail.com
