## Сбор данных о файлах в папке

### Подготовка

- Не требуется

### Запуск

Пример запуска без выделения текста

```

powershell.exe -ExecutionPolicy Bypass -Command "./explore-folder.ps1" -folder "c:\\work\\test" -outfilename folder_test

```

Пример запуска с выделением текста

```

powershell.exe -ExecutionPolicy Bypass -Command "./explore-folder.ps1" -folder "c:\\work\\test" -outfilename folder_test -extruct

```

Пример запуска без выделения текста сбора данных о всех папках общего доступа, с компьютеров зарегистрированных в указанной организационной единице

```

powershell.exe -ExecutionPolicy Bypass -Command "./explore-folder.ps1" -base DC=acme``,DC=local -server dc.acme.local -outfilename folder_test

```



Параметры:

| Имя         | Назначение                                                                   |
|-------------|------------------------------------------------------------------------------|
| folder      | Корневая папка(локальная или сетевая) для сбора данных                       |
| base        | [Необязательный] Корневая OU для зачитывания списка компьтеров из ActiveDirectory         |
| server      | [Необязательный] Имя домен-контроллера для зачитывания списка компьтеров из ActiveDirectory                          |
| user        | [Необязательный] Имя пользователя под которым производится запрос. Если не заданно, то выводится диалог с запросом |
| pwd         | [Необязательный] Имя пользователя под которым производится запрос. Если не заданно, то выводится диалог с запросом |
| outfilename | Имя файла результатов                                                        |
| extruct     | Выделять ли текст из doc, docx, xls, xlsx                                    |

