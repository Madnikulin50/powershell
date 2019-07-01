## [IN DEVELOPMENT] Экспорт списка почтовых ящиков из Exchange

### Подготовка

Запуск нужно осуществлять из Windows PowerShell ISE с включенным Exchange Management Shell


### Запуск

Пример запуска
```
powershell.exe -ExecutionPolicy Bypass -Command "./export_mb_folders.ps1" -outfile folders.csv
```
Параметры:

| Имя     | Назначение                                      |
|---------|-------------------------------------------------|
| outfile | Имя файла результатов                           |


