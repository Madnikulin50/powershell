## Сбор данных о сетевых папках в домене

### Подготовка

- Не требуется

### Запуск

Пример запуска

```

powershell.exe -ExecutionPolicy Bypass -Command "./explore-folder.ps1" -folder "c:\\work\\test" -outfilename folder_test

```
Параметры:

| Имя         | Назначение                                      |
|-------------|-------------------------------------------------|
| folder      | Корневая OU для экспорта                        |
| outfilename | Имя файла результатов                           |
