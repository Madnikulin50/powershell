## Сбор данных о файлах в папке

## Сбор событий EventLog c удаленных рабочих станций

### Требования для использования
+ Операционная система Windows 7+, Windows 2012+. Рекомендуемая Windows 10x64.1803+, Windows 2019x64
+ Windows PowerShell 5+, Рекомендуется Windows PowerShell 5.1
+ Remote Server Administration Tools for Windows 10 (или другой для соответвующей версии ОС)
+ Права на чтение файлов из инспектируемых файлов и папок
+ + Права на чтение данных из ActiveDirectory (Read all user information) [Дополнительно](https://social.technet.microsoft.com/Forums/en-US/c8b5886a-f0f1-4e20-b083-d36521d4dec6/delegation-to-read-all-users-properties-in-the-domain?forum=winserverDS)


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
| folder      | [Необязательный] Корневая папка(локальная или сетевая) для сбора данных                       |
| base        | [Необязательный] Корневая OU для зачитывания списка компьтеров из ActiveDirectory         |
| server      | [Необязательный] Имя домен-контроллера для зачитывания списка компьтеров из ActiveDirectory                          |
| user        | [Необязательный] Имя пользователя под которым производится запрос. Если не заданно, то выводится диалог с запросом |
| pwd         | [Необязательный] Имя пользователя под которым производится запрос. Если не заданно, то выводится диалог с запросом |
| outfilename | Имя файла результатов                                                        |
| extruct     | Выделять ли текст из doc, docx, xls, xlsx                                    |

