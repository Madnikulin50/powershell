## Экспорт пользователей из AD

### Подготовка

- Перед запуском должен быть установлен Remote Server Administration Tools for Windows 10 (или другой для соответвующей версии ОС)

### Запуск

Пример запуска

```

powershell.exe -ExecutionPolicy Bypass -Command "./export_ad_users.ps1" -base DC=testdomain`,DC=local -server kappa.testdomain.local -outfilename export_users

```
Параметры:

| Имя         | Назначение                                      |
|-------------|-------------------------------------------------|
| base        | Корневая OU для экспорта                        |
| server      | Имя домен-контроллера                           |
| outfilename | Имя файла результатов                           |

После запуска будет выведено окно логина на домен-контроллер, нужно ввести логин-пароль пользователя имеющего право читать учетные данные из домена
