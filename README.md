# Practic-428-1
Студент группы 428/1 Ислентьев Н.В.    Приложение по теме: "генератор заголовков"

Темой практики было выбрано: "Создание документа из хедера" 
Приложение было реализовано на языке программирования Python, основными библиотеками выбраны: tkinter, os, docx и reportlab.
Средой разработки был выбран Intellij IDEA Community Edition 2022.2. 
Приложение реализовано в виде исполняемого файла формата .exe, последствием открытия которого, пользователю становится доступен графический интерфейс(GUI), открывается в небольшом окне. 

В самом приложении предусмотрены следующие функции:
1. Ввод заголовков(каждый с новой строки)
2. Выбор формата(предусмотрены 2 формата: pdf и docx
3. Имя генерируемого файла(здесь вводится полное имя файла)
4. Обзор(выбор места сохранения файла)
5. Выбор шрифта(из-за того, что пришлось вручную прописывать шрифты, их всего 3: Arial, TNR и Courier
6. Выбор размера шрифта(на выбор предоставляются шрифты от 10 до 18)
7. Кнопка генерации документа(выполняет свою функцию :D)

Небольшая проблемка при создании: P.S. Ошибка в основном со шрифтами!
1. Первой, и по-совместительству, главной ошибкой стали шрифты. При генерации файла формата pdf, если пользователь вводил заголовки на русском языке, вместо корректного отображения выводились чёрные квадраты.
Данная ошибка была пофикшена посредством установки шрифтов из интернета, потому что по какой-то неизвестной причине, шрифты отказывались работать, если их устанавливать через консоль разрботчика.
2. Второй проблемой стало то, что при генерации документа формата docx, все все заголовки генерировались по умолчанию шрифтом calibri. Чуть позже стало понятно, что это связано с тем, что python-docx устанавливает calibri по умолчанию...
Исправлено посредством написания корректной функции выбора шрифта в коде.

Т.к. приложение реализовано исполнительным файлом exe, его и прикрепляю, также добавлю код приложения.
