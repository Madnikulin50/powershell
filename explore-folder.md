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


Параметры:

| Имя         | Назначение                                                                   |
|-------------|------------------------------------------------------------------------------|
| folder      | Корневая папка(локальная или сетевая) для сбора данных                       |
| outfilename | Имя файла результатов                                                        |
| extruct     | Выделять ли текст из doc, docx, xls, xlsx                                    |

