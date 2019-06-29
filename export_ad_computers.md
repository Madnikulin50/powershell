## Экспорт компьютеров из AD

### Подготовка

- Перед запуском должен быть установлен Remote Server Administration Tools for Windows 10 (или другой для соответвующей версии ОС)


### Запуск

Пример запуска
```
powershell.exe -ExecutionPolicy Bypass -Command "./export_ad_computers.ps1" -base DC=testdomain`,DC=local -server kappa.testdomain.local -outfile export_computers.csv

```
Параметры:

| Имя     | Назначение                                      |
|---------|-------------------------------------------------|
| base    | Корневая OU для экспорта                        |
| server  | Имя домен-контроллера                           |
| outfile | Имя файла результатов                           |

После запуска будет выведено окно логина на домен-контроллер, нужно ввести логин-пароль пользователя имеющего право читать учетные данные из домена
