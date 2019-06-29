# Скрипты для наполнения базы объетов Makves IRP

## Подготовка к работе

1. Убедиться что установлен Windows Management Framework 5.0 или 5.1
2. Убедиться что выключены все галочки в Remote Server Administration Tools (RSAT). См. https://blogs.msdn.microsoft.com/adpowershell/2009/03/24/active-directory-powershell-installation-using-rsat-on-windows-7/


## Организация репозитория
У каждого скрипта есть свой файл описания с тем же именем и расширением md

## Список скриптов

| Имя                       | Назначение                                      |
|---------------------------|-------------------------------------------------|
| export_ad_users           | Экспорт информации о пользователях из AD        |
| export_ad_computers       | Экспорт информации о компьютерах из AD          |
| export_mb_folders         | Экспорт почтовых ящиков из Exchnge              |
| explore_share             | Сбор информации о файлах                        |
