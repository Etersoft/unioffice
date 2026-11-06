# UniOffice: сборка и проверка Wine/MinGW

Рабочие заметки от 2026-11-05. Здесь разделены проверенные результаты,
изменения исходников и ещё не завершённые проверки. Успешная компиляция
не означает, что автоматизация Excel или Word уже работает.

## Актуальное состояние после уточнений

Пользователь уточнил: нужны только Windows PE DLL. Сборка `.dll.so` была
ошибочным отклонением и не является поддерживаемым вариантом.

- winegcc запускать на **builder64**: `make libraries TOOLCHAIN=winegcc -j4`.
  Выбран `winegcc -b x86_64-windows`: Clang/lld, Wine headers/import libraries,
  результат — PE32+ x86-64. MinGW на этом хосте для данного варианта не нужен.
- На **Fedora только прямой MinGW**:
  `make libraries TOOLCHAIN=mingw MINGW=i686-w64-mingw32 -j4` (PE32 i386).
- Оба хоста видят общий `/srv` по NFS. Оба варианта успешно собраны после
  последних изменений. COM smoke-тест для обоих PE DLL: **0 ошибок**.
  Проверены экспорты, фабрика, LockServer/CanUnloadNow, чтение TLB, выгрузка.
- PE64 тест собран на builder64:
  `winegcc -b x86_64-windows -o unioffice_excel/build/winegcc-64/com-smoke-pe.exe tests/com-smoke.c -lole32 -loleaut32 -luuid`.
  PE32 тест собран на Fedora через `i686-w64-mingw32-gcc` с теми же библиотеками.
  Исполнялись локальным `/usr/lib/wine-etersoft/wine64` в изолированном префиксе;
  пути DLL и TLB передавались как абсолютные `Z:/srv/...`.
- В smoke-тесте убран CoInitialize: проверяются непосредственные DLL exports
  и TLB, создание COM backend не требуется. Это исключает запуск RPC-служб.
- Clang выявил ошибки старого кода: прямой доступ к VARIANT заменён макросами
  `V_INT`/`V_DISPATCH`; исправлены небезопасные вызовы swprintf, включая запись
  строки в себя. Для шести возвращающих VARIANT методов I_Range добавлены
  адаптеры к сигнатурам WIDL с явным буфером результата после This.
- Удалён отладочный fprintf из DllMain: после этого исчезло зависание ELF
  PROCESS_ATTACH. Причина внутри CRT не установлена. Каждый ti_* теперь
  освобождается через собственную vtable.
- Запрос задачи Wine про argc в старом ELF-режиме был отправлен на согласование
  (confirmation_id `29afc298-d800-49a2-bdfb-f560fe630e2d`), затем отозван после
  уточнения цели. Это не дефект целевой PE-сборки; новая задача не создавалась.
- Astra пока не требуется. Сообщить пользователю, если сложность разбора
  потребует её подключения.

Ниже сохранены исторические эксперименты, включая неактуальную схему ELF.
При противоречии руководствоваться этим разделом и README.

## Проверка MSI32

Добавлен `scripts/build-msi.py`; `MSI/Makefile` вызывает его. `make` теперь
собирает библиотеки; установщик собирается явно. PE64 отклоняется до запуска
Wine: исходные Component/Directory/шаблон рассчитаны на 32 бита.

MSI32 успешно упакован на builder64 из MinGW-артефактов Fedora в
`MSI/build/mingw-32/unioffice.msi`. Существовавший `MSI/unioffice.msi` сохранён.
Рабочие таблицы и CAB создаются во временном каталоге `~/tmp/claude`;
размеры в File.idt вычисляются из реальных DLL/TLB. Binary.idt импортируется
вместе с каталогом Binary/ потоков, повторный импорт самих .ibd не нужен.

Используется обычный Wine вместо отсутствующего wine98; выбранный через WINE
loader запускает также winepath.exe. MsiDb — GUI-программа, поэтому на сервере
нужен Xvfb. Логи/тестовый префикс: `~/tmp/claude/unioffice-com-check/prefix`.

Подтверждён дефект Wine: импорт Control.idt с документированным префиксом
1251 в третьей строке завершается ошибкой. Аналогичный формат есть в Property
и RadioButton. Обход применяется только к временным копиям: убрать префиксы,
импортировать _ForceCodepage, задать LC_ALL=ru_RU.UTF-8 дочерним процессам.
Один _ForceCodepage недостаточен: исходный IDT читается через CP_ACP, и ACP1252
искажает русский текст. Проверка tests/msi-smoke.c через MsiRecordGetStringW
подтвердила точное сохранение ОАО "Etersoft" после обхода (чтение при ACP1252).
Текст лицензии целиком сохранён; его размер не был причиной сбоя импорта.

Формат подтверждён [документацией Microsoft](https://learn.microsoft.com/en-us/windows/win32/msi/archive-file-format).
Задача Wine создана: [19875](https://bugs.etersoft.ru/show_bug.cgi?id=19875),
компонент «Установка программ / MSI».

Проверено административное извлечение (`msiexec /a /qn`) в предварительно
созданный каталог `~/tmp/claude/unioffice-msi-extract`, с Windows-путями через
обратную косую черту. SHA-256 извлечённых DLL/TLB совпали с исходными:

- DLL: `a561408665a7ccbc3516233be90434407207e1419d57396ec43bb6284c05176a`;
- TLB: `6f4ec70a924f82a9374e974ec15137d84c7b461d1a5b81da8446db14c39ebd24`.

Добавлен /s и кавычки вокруг пути DLL в Windows custom actions регистрации
и удаления. Тихая установка MSI32 в тестовом префиксе завершилась успешно,
regsvr32 зарегистрировал C:\windows\syswow64\unioffice_excel.dll; ключ
CLSID\{00024500-0000-0000-C000-000000000046}\InprocServer32 проверен в /reg:32.
Тихое удаление (`msiexec /x {643315BA-DEF8-4CBE-AB26-C54B71535DFD} /qn`)
также прошло успешно: regsvr32 сообщил об отмене регистрации, DLL удалена.
Журналы установки и удаления: `~/tmp/claude/unioffice-msi-{install,uninstall}.log`.

В тестовом префиксе builder64 ключ HKCR\com.sun.star.ServiceManager\CLSID
отсутствует. Наличие Linux LibreOffice (`/bin/soffice`, `/usr/lib64/libreoffice`)
не означает наличие Windows COM backend. Полные VBS-тесты пока не запускались.

## Windows OpenOffice backend: проверка 4.1.17

На builder64 скачан официальный Windows x86 Apache OpenOffice 4.1.17 с
SourceForge (ссылка с сайта Apache). SHA-256 проверен по файлу с downloads.apache.org:
`sha256sum -c installer.sha256` → OK. Файлы находятся в
`~/tmp/claude/unioffice-aoo/`, распаковка через `7z x -oinstaller`.

Создан отдельный `WINEPREFIX=~/tmp/claude/unioffice-aoo/prefix`.
Установка через `/usr/lib/wine-etersoft/wine64 msiexec /i <openoffice4117.msi>
/qn /norestart ADDLOCAL=ALL SETUP_USED=1` под Xvfb завершилась с кодом 0.
Без SETUP_USED=1 MSI отказывает в LaunchConditions: требует запуск через setup.exe.
Это условие установщика, а не ошибка Wine. Путь установленного офиса:
`C:\Program Files (x86)\OpenOffice 4`. Log: `unioffice-aoo/install.log`.

MSI зарегистрировал ServiceManager и CLSID
`{82154420-0FBF-11d4-8313-005004526AB4}`, LocalServer32 указывает на soffice.exe.
Первый быстрый запуск COM до инициализации профиля не прошёл; повторный после
запуска офиса прошёл полностью. Проверка `tests/openoffice-backend.vbs`:
ServiceManager → Desktop → новый Calc → запись/чтение ячейки → закрытие.
Скрипт `tests/run-openoffice-backend.sh` запускается через `sh` под Xvfb;
Wine loader задаётся WINE. ServiceManager получает до 30 попыток ожидания.
В конце освобождаются COM-ссылки и вызывается Desktop.terminate.

В этом же префиксе установлен наш MSI32, затем прошла проба
`tests/excel-backend.vbs`: Excel.Application, Workbooks.Add, запись и чтение
Cells(1,1).Value, закрытие книги и Excel.Quit. Итог: UniOffice integration: PASS.
Команда запуска: `sh tests/run-openoffice-backend.sh excel-backend.vbs
"UniOffice integration: PASS"` в том же WINEPREFIX/WINE под Xvfb.

Важно: Wine cscript вернул 0 даже при VBScript runtime error. Проверяющий shell
скрипт проверяет и exit status, и явную строку PASS; один exit status недостаточен.

Первый запуск исходной асинхронной цепочки в
`~/tmp/claude/unioffice-excel-suite-5hr5n_7l` не завершился за 150 секунд.
Он смешивал 32-битный первый cscript и последующие 64-битные cscript.exe,
получая каскадные ошибки. Это не достоверная оценка совместимости методов.

`scripts/run-excel-suite.py` переделан: каждый из семи исходных тестов копируется
в отдельный временный подкаталог, сохраняет проверки и CP1251, получает только
замену MsgBox на файл завершения. Отдельные каталоги предотвращают автоматический
запуск следующего скрипта старой цепочкой. Каждый тест запускается через явный
32-битный cscript после старта headless OpenOffice. Предел — 60 секунд на тест;
после timeout прогон прекращается, чтобы зависший COM-вызов не влиял на следующие.
Рабочий каталог повторного прогона:
`~/tmp/claude/unioffice-excel-suite-qp5hmjge`. Все семь тестов завершились:
по исходным счётчикам 134 успеха и 31 ошибка. №5–7 прошли, №1–4 имеют ошибки.
У №6 12 успехов в счётчике, но 11 SUCCESS-маркеров. Подробности:
[OpenOffice 4.1.17 results](openoffice-4.1.17-results.md).

Исторические записи ниже об отсутствии backend относятся к прежнему префиксу
`unioffice-com-check`, а не к новому `unioffice-aoo/prefix`.

## Исправление PageSetup

Воспроизведение с диагностическим `unioffice_excel.log` выявило общий отказ:
доступ к `StyleFamilies` через DISPATCH_PROPERTYGET. UNO предоставляет метод
[getStyleFamilies](https://www.openoffice.org/api/docs/common/ref/com/sun/star/style/XStyleFamiliesSupplier.html).
Все 24 вызова в i_pagesetup.c переведены на DISPATCH_METHOD/getStyleFamilies.
MinGW на Fedora и winegcc PE64 на builder64 успешно собраны.
В тестовом префиксе проверена обновлённая MinGW DLL (установленная DLL заменена
из результата сборки; после регрессии MSI32 пересобран с исправлением).

Выборочный прогон `scripts/run-excel-suite.py --tests 4`:
`~/tmp/claude/unioffice-excel-suite-6bcuesbf`, счётчики 34/0, маркеры 33/0.
Legacy-счётчик снова расходится с числом маркеров; ошибок нет.
Полный регрессионный прогон: `~/tmp/claude/unioffice-excel-suite-_8c6q9m1`.
Все семь тестов завершились: №4–7 проходят, осталось 7 прежних ошибок
в №1–3 (шрифты 4, Outline 1, Border 2). Новых ошибок не обнаружено.

## Исправление Subscript/Superscript (2026-11-06)

Getters были E_NOTIMPL. Setters записывали BOOL в VARIANT с тегом VT_I4,
а UNO CharEscapement требует signed short. Но исправления типов недостаточно:
Calc не предоставляет CharEscapement у диапазона. Реализация теперь получает
ячейки через getCellByPosition, создаёт текстовый курсор, выделяет текст через
gotoStart/gotoEnd и читает/записывает VT_I2 CharEscapement курсора.

Изменяются все ячейки диапазона; getter возвращает Null при разных состояниях.
Входной VARIANT копируется через VariantCopyInd и преобразуется без изменения
переданного значения. Новая проверка tests/font-script.vbs покрывает True/False,
передачу переменной, числовое -1, переключение между индексами, диапазон 2x2 и Null.
Она завершилась Font script integration: PASS. Выборочный исходный test_1:
`~/tmp/claude/unioffice-excel-suite-wo8_d87e`, 32 SUCCESS / 0 FAILED.
Обе сборки проходят (MinGW/Fedora и winegcc/builder64).

При повторной загрузке исправленной DLL один расширенный прогон видел прежнее
поведение; после остановки только тестового wineserver проблема исчезла.
Причина кеширования/оставшихся процессов не установлена. Для следующих замен
DLL использовать атомарную замену или остановить процессы тестового префикса,
не перезаписывать используемый файл. Исходные непройденные проверки шрифта сняты;
остаются Outline.ShowLevels и два свойства Border.

## Исправление Outline.ShowLevels (2026-11-06)

В IDL выходной VARIANT не был retval. Automation требовал третий аргумент и
отказывал до входа в реализацию. Исправлено на [out, retval]. Метод обрабатывает
обе оси, пропущенный/нулевой аргумент оставляет соответствующую ось без изменений,
отрицательные уровни отклоняются. Excel level 1 соответствует Calc depth 0.

Новый tests/outline-levels.vbs проверил результат вызова и реальные IsVisible
строк/столбцов через прямой UNO: сворачивание обеих осей, раскрытие, нули и
пропущенный первый аргумент. Все проверки PASS. Исходный test_2: 13/0,
`~/tmp/claude/unioffice-excel-suite-l2f44geo`. Обе сборки проверены; после изменения
IDL в префиксе заменены DLL и TLB.

Для устойчивого тестового запуска добавлен tests/openoffice-ready.vbs:
запускатели ждут реальной COM-активации ServiceManager, а не только sleep.
Неудачный запуск до готовности офиса не считается отказом тестируемого метода.
Шрифтовое исправление сохранено в 1624f69. Остались две проверки границ.

## Диагональные границы и итоговая регрессия (2026-11-06)

Вызов dynamic PROPERTYGET для DiagonalTLBR/DiagonalBLTR из UniOffice мог
завершаться ошибкой, хотя прямой доступ к тому же range из VBScript работал.
Восемь мест чтения диагональных структур переведены на явный UNO
getPropertyValue. Исходная причина различия внутри COM bridge не установлена;
ошибка Wine на основании этого наблюдения не объявлялась.

Новый tests/border-diagonal.vbs проверяет не только отсутствие HRESULT ошибки,
но и round trip LineStyle/ColorIndex по всем 12 legacy indices, реальные color/
width в UNO для обеих диагоналей, double и удаление линии. Все проверки PASS.
Исходный test3: 30/0. Обе сборки прошли. В tools.c отдельно инициализирован
EXCEPINFO и добавлен HRESULT/flags/DISPID в диагностику (bd464bb).

Полный прогон `~/tmp/claude/unioffice-excel-suite-dh6epjsc`: все семь тестов
завершились, return code 0; исходные счётчики 166/0, SUCCESS-маркеров 164
(известные расхождения test4 и test6). Полная таблица в results-документе.
Итоговый MSI32 пересобран из актуальных DLL/TLB. Unicode metadata check: 0.
В чистом тестовом префиксе unioffice-com-check административное извлечение прошло,
cmp подтвердил побайтовое совпадение DLL и TLB с протестированными артефактами.
Извлечённый payload: ~/tmp/claude/unioffice-final-msi-extract.
Тот же /a запуск в префиксе unioffice-aoo с установленным продуктом остановился
на InstallAdminPackage. Причина различия префиксов не изолирована; это пока
наблюдение для отдельного разбора, а не доказанный дефект содержимого MSI/Wine.

Шрифтовое и Outline-исправления сохранены в 1624f69 и 63d39c4, readiness —
96630e2. Вопрос Borders.Count=12/legacy aliases остался вне этой правки: тесты
его семантику не проверяют. Полная совместимость PE64 automation не заявляется.

Создан отдельный репозиторий эксперта `/home/lav/Projects/Claude/win-builder`
(canonical `/srv/lav/Projects/Claude/win-builder`), начальный commit 7100a78.
После итоговой регрессии туда переносятся подтверждённые выводы.

## PE64 Automation и unregister typelib (2026-11-06)

Создан отдельный prefix ~/tmp/claude/unioffice-aoo64/prefix из остановленной
копии PE32/OpenOffice окружения. PE64 DLL/TLB зарегистрированы в system32,
используется system32/cscript.exe. Windows OpenOffice остаётся x86; COM
marshaling между архитектурами работает в проверенных сценариях.

Все семь тестов через winegcc PE64 прошли: 166/0 counters, 164 SUCCESS markers,
каталог ~/tmp/claude/unioffice-excel-suite-8qgvlq1p. Прямой OpenOffice probe,
Excel.Application round trip и расширенные Font/Outline/Border probes тоже PASS.
Запускатели получили CLIENT_BITS=64 / --bits 64; default32 сохранён. Изменения
сохранены в 1ff5ca9. Результаты и команды: [PE64](pe64-results.md).

Проверка регистрации выявила дефект UniOffice: PE64 UnRegisterTypeLib с
SYS_WIN32 удалял win32 и сохранял win64, несмотря на success. Native test
воспроизвёл 2 failure(s). В ce5d3fd добавлен выбор SYS_WIN64 для _WIN64.
После исправления tests/typelib-registration.c даёт 0 failures на обеих
архитектурах, сохраняет соседнюю ветку и восстанавливает собственную регистрацию.
После цикла registration/unregistration исходный test2 прошёл 13/0 и через
PE64 (unioffice-excel-suite-ldpjs0yw), и через PE32 (unioffice-excel-suite-0j9voquh).
MSI32 пересобран из актуальных артефактов; Unicode metadata check: 0.
MSI64 и Word Automation по-прежнему не заявлены; исторические оговорки ниже
об отсутствии PE64 runtime-проверки относятся к прежним этапам.

## MSI64 и совместное удаление (2026-11-06)

Добавлена упаковка AMD64 на основе временных архитектурных таблиц: Summary
x64;1033, schema200, Component.Attributes258, System64Folder, отдельные Product/
Upgrade/Component GUID. Для каждого билда новый PackageCode. MSI32 сохранён,
пакеты проверены расширенным metadata/Unicode тестом: 0 failures каждый.

Совместная установка выявила потерю общей ProgID при удалении одного продукта.
Исправление 85a457b сохраняет общие классы, пока есть peer InprocServer32 в
соседнем registry view. Native regression red→green на обеих архитектурах.
Условия MSI custom actions переведены на component action states, чтобы repair
не снимал регистрацию без восстановления.

scripts/test-msi-coexistence.py подтвердил обе архитектуры, repair MSI64,
оба порядка удаления с работающим оставшимся клиентом и очистку после последнего
удаления. MSI coexistence: PASS. Prefix ~/tmp/claude/unioffice-msi64/prefix,
журналы coexistence-logs/. В нём оба продукта в конце удалены; прежние рабочие
префиксы не затронуты. Подробности и GUID: [MSI64](msi64-results.md).
Исторические записи о неподдерживаемом MSI64 ниже теперь относятся к прежним этапам.

## Независимые Borders enumerators (2026-11-06)

После MSI64 разобран узкий дефект коллекции Borders: общий cursor ломал
вложенный For Each, Skip повторял последний элемент и переполнялся на ULONG_MAX,
Clone возвращал E_NOTIMPL. Новый script воспроизвёл отказ первого внутреннего
цикла; native probe подтвердил ошибки cursor/Skip/Clone на прежней DLL.

GetEnumerator и Clone теперь создают независимые контексты с удержанием
range/backend; Skip ограничивает позицию концом без сложения больших чисел.
Native test и вложенный/повторный VBS обход прошли PE32 и PE64; diagonal probe
PASS на обеих, legacy test3 — 30/0, markers30/0. Count12/aliases не менялись;
Word по-прежнему заготовка. Остальные legacy tests в этом этапе не повторялись.
Подробности: [Borders enumeration](borders-enumerator-results.md).

## Восемь разных Borders и проверка индексов (2026-11-06)

Count/enumeration теперь основаны на восьми разных XlBordersIndex keys5..12,
а не двенадцати ключах с четырьмя дублями. Aliases1..4 оставлены для lookup,
не выдаются отдельно в обходе. Microsoft перечисляет восемь ключей, но общее
описание говорит о четырёх границах; LibreOffice VBA реализует восемь. Exact
Excel order и native runtime Count не сравнивались; выбран ascending5..12.

Invalid Item/default keys0/-1/13/INT_MIN/INT_MAX теперь дают DISP_E_BADINDEX и
NULL output. Native regression до исправления дал21 failures. Новый VBS
проверяет уникальность через разные цвета, aliases, массовые Color/LineStyle,
диагонали, вложенные/повторные обходы. Полные результаты и ограничения:
[Borders membership](borders-membership-results.md).

При расширении проверки найден отдельный старый stub Border.Weight:
put возвращает S_OK без изменений, get — E_NOTIMPL. Отдельный известный failing
probe tests/border-weight.vbs сохранён; поддержку Weight ещё предстоит реализовать
и проверить через реальные UNO widths. Новых подтверждённых Wine defects нет.

## Реализация Border.Weight (2026-11-06)

Заглушки get/put Weight заменены реальными UNO widths. Общие helpers для
LineStyle/Weight убрали дублированный код: single stroke записывается в Outer,
double использует обе линии и gap, Color сохраняется. Четыре категории имеют
номинальные total widths2/26/88/141 сотых мм; double делит total на три части,
getter учитывает округление. Это проверенная политика адаптера, не native Excel
rendering oracle. LineStyleNone стирает размеры; Weight делает линию видимой.

TableBorder setters включают только выбранный Is...LineValid: смешанные соседние
границы не перезаписываются. Aggregate Weight проходит шесть недиагональных
ключей и возвращает ошибки child calls. Getter пока сохраняет чтение верхней
границы: legacy enum retval не позволяет Null для смешанной aggregate Weight.

Расширенный border-weight.vbs дал374 OK на PE32 и PE64: все четыре веса,
12 accepted indices, single/double, прямые UNO dimensions, aliases, aggregate,
invalid weights и сохранение разных A1/A2 edges. Native COM test проверил NULL
outputs/invalid enums; enumeration и diagonal probes также PASS.
Полный набор7 tests прошёл на обеих архитектурах: legacy158/0, markers156/0.
Старые расхождения counters tests4/6 сохраняются. Results:
~/tmp/claude/unioffice-excel-suite-mw8b7s1v (32), -75jv9aff (64).
Оба MSI обновлены, metadata/Unicode0 каждый; lifecycle не повторялся.
Подробности и ограничения: [Weight results](border-weight-results.md).
Новых подтверждённых дефектов Wine нет; нативная Windows не проверялась.

## Mixed Borders.Weight и сохранение ABI (2026-11-06)

Automation getter теперь даёт VT_I4 для одинакового веса и VT_NULL для разных.
Legacy I_Borders IDL/vtable/getter XlBorderWeight* сохранены без изменения IID:
нативные клиенты по-прежнему читают top edge. Добавлен pure dispinterface
D_Borders (38CD5803-18E6-4F67-9525-110B4F31896B), Weight metadata=VARIANT;
IDispatch обслуживает mixed getter отдельно, остальные calls идут через старую
typeinfo/vtable. TypeInfo AddRef/outputs проверены native test с enum-buffer canaries.

Сравниваются шесть недиагональных сторон с пропуском отсутствующих inner lines
у single cell/row/column. Неоднозначные UNO validity flags требуют per-cell
проверки нужных сегментов, чтобы разные colors/styles при одинаковом Weight
не давали ложный Null. Mixed probe55 checks и физический Weight probe374 прошли
на PE32 и PE64. Полные7 legacy tests: counters158/0, markers156/0 каждый.
Results: ~/tmp/claude/unioffice-excel-suite-w_qyjlm2 (32), -hjmdhjh4 (64).
Первый параллельный PE32 suite-e85hajwi упал в test5 после создания workbook,
до Weight. Отдельный5/6/7 suite-wwimq370 прошёл, затем прошёл полный повтор32.
Первый отказ не скрыт, root cause не установлен и Wine defect не объявлялся.

Дополнительно выявлена порча строкового argument Application.Range на PE64:
отдельный Range/Borders.Count probe воспроизвёл её без Weight. Копирование
аргументов VariantCopyInd в owned locals устранило ошибку; оба final probes
прошли32/64. Чистые вложенные VBScript arrays проходили без UniOffice.
Гипотезы о primitive-result assignment/VariantInit не устраняли симптом;
новый Wine defect не доказан. Исправление Range сохранено отдельным коммитом9922be8.

DLL/TLB установлены парами; оба MSI пересобраны, metadata/Unicode0. После
установки каждого mixed probePASS. Lifecycle/coexistence прошёл repair64 и оба
порядка удаления с работающим surviving client, затем очистку последнего.
Только prefix unioffice-msi64 в конце без обоих продуктов; рабочие fixtures сохранены.
Подробности: [mixed Weight](border-weight-mixed-results.md),
[Range arguments](range-argument-results.md). Native Excel/Windows не проверялись.

## Возможности backend и ошибки aggregate LineStyle (2026-11-06)

Проверка Windows AOO4.1.17 через32/64 script hosts: TableBorder/DiagonalTLBR
есть, TableBorder2/LeftBorder2/DiagonalTLBR2 нет. Reflection знает BorderLine,
но не BorderLine2/BorderLineStyle; у классической линии LineStyle даёт1B6.
Расширенный struct документирован LibreOffice3.4 и не является подтверждённой
возможностью установленного Apache backend. Пунктирная отрисовка не реализована.

Найден конкретный дефект UniOffice: отдельная граница возвращает E_NOTIMPL,
а aggregate LineStyle глотает HRESULT и возвращает S_OK; Value наследует ошибку.
Новый VBS reproducer упал на первом aggregate Dash прежней DLL. Setter теперь
возвращает первый child failure с Release; getter проверяет NULL и возвращает
реальный HRESULT. Invalid style12345 даёт E_INVALIDARG; ABI/IDL не менялись.

На PE32/MinGW Fedora и PE64/winegcc builder64 прошли style regression, native
HRESULT/null/ABI checks, Weight374, mixed55, диагонали и legacy test3 (22/0,
markers22/0). Остальные6 tests не повторялись. Outputs:
~/tmp/claude/unioffice-excel-suite-6wh01dyo (32), -e0zfiu8l (64).
Оба MSI пересобраны, metadata/Unicode0; installer lifecycle не повторялся.
Подробности: [Border styles](border-style-results.md). Новых Wine defects нет.

## Worksheet.Range/Cells: входные аргументы и parent lifetime (2026-11-06)

Аудит после Application.Range воспроизвёл ещё две порчи аргументов: PE64
Worksheet.Range(address) менял строку после Borders.Count; PE32 Cells(row,"3")
проходил первый string check и падал на втором. Worksheet.Range/default indexer
теперь копируют входы VariantCopyInd в owned locals и освобождают их на всех
выходах. Индексы преобразуются только в собственных копиях; conversion failure
возвращается до создания cell. Application.Cells get/put и Worksheet.Cells put
теперь возвращают HRESULT child calls и не используют несуществующий результат.

Endpoints Worksheet.Range разрешаются независимо как BSTR или Range через QI;
поддержаны string/object combinations, invalid types отвергаются. Temporary
endpoint interfaces и Cells collection освобождаются. Native public ABI/IDL не менялись.

При промежуточной PE64 проверке Cells упал до assertions. Отдельный Parent probe
без строковых индексов воспроизвёл execute fault по адресу2 и timeout124. Причина:
Worksheet.Cells отдавал private cache с weak pwsheet, Application.Cells освобождал
временный Worksheet до использования Range. Commit2ad7996 возвращает отдельный
Range с strong worksheet reference; private weak cache остаётся внутренним.
Прежний разовый legacy test5 crash без matching backtrace ему не приписывается.

На PE32/MinGW Fedora и PE64/winegcc builder64 прошли probes40/37/49 checks
(Worksheet.Range inputs / Cells inputs / parent lifetime). Native тест проверил
E_POINTER, E_INVALIDARG, DISP_E_TYPEMISMATCH, NULL failed outputs и intact BSTR:
0 failures на обеих. Старые Application.Range/mixed Weight probes также PASS.
Финальный полный прогон7 tests: counters158/0, markers156/0 каждый; старое
расхождение tests4/6 сохранено. Outputs ~/tmp/claude/unioffice-excel-suite-kndhu5eq
(32), -e3lqpp45 (64). Все исходные отказы сохранены. Оба MSI пересобраны,
metadata/Unicode0; installer lifecycle не повторялся, регистрация/IDL не менялись.
Подробности: [Range/Cells](range-cell-ownership-results.md),
[Cells parent](cells-parent-lifetime-results.md). Нативная Windows не проверена;
новых подтверждённых Wine defects нет. Range.Item stub остаётся отдельной задачей.

## Range.Item и индексированное присваивание (2026-11-06)

Range.Item больше не заглушка: cell ranges поддерживают положительные
относительные row/column, буквенные столбцы, linear row-major index и assignment.
Numeric default indexer и Range.Cells используют ту же выборку. Публичные
IDL/IID/vtable сохранены: обычный LCID argument из старого IDL скрывается от
скриптов в Invoke. Legacy BSTR lookup для Worksheet.Range сохранён отдельно.
Rows/Columns collection rules, multi-area и nonpositive offsets не покрыты.

UNO не разрешил Item за пределами исходного subrange: теперь offsets переводятся
в абсолютные sheet coordinates, child создаётся через owned worksheet Cells.
Входы и RHS копируются VariantCopyInd; cleanup освобождает все temporaries.

Новый C-тест выявил9 дальнейших Weight failures. Контроль без range calls прошёл,
без value writes всё ещё падал. Прямой C PROPERTYGET TableBorder вернул80004002,
а getPropertyValue("TableBorder") на том же target — S_OK. Прямой UNO VBS control
прошёл оба пути; Wine defect этим не установлен. TableBorder read/write переведены
на явные UNO методы, включая Color/Weight. Тест покрывает Weight/Color/LineStyle
после Item, точные HRESULT, byref Double, BSTR ownership, overflow и bad argcount.

Одна быстрая NFS-правка не привела к incremental rebuild на Fedora: i_range.o
оказался старше исходника, несмотря на совпадающие source hashes. Исправлено
targeted make -W i_range.c; перед повтором сверены deployed DLL. Это не runtime bug.
Подробности и команды: [Range.Item](range-item-results.md).

Финальные PE32/PE64: Item36 checks PASS, native COM0 failures, прежние
Worksheet.Range/Cells/parent/Application.Range и mixed Weight/Weight/style probes
PASS. Полные7 legacy tests каждый: counters158/0, markers156/0; outputs
~/tmp/claude/unioffice-excel-suite-clxbrb0m (32), -f807h133 (64). Оба MSI пересобраны,
metadata/Unicode0; lifecycle не повторялся (IDL/registration unchanged). DLL в
fixtures побайтно совпадают с build outputs. Native Windows не проверялась.

## Application/Worksheet.Cells: общий dispatch (2026-11-06)

Отдельный cells-indexing.vbs воспроизвёл9 failures в39 выполненных checks на
PE32 и PE64: Application.Cells(4), одноиндексный put обоих owners, неверный
target при Cells("2","AA") put и отсутствие ошибки для invalid string index.
Причины — VT_EMPTY вместо omitted column, требование двух индексов для put и
legacy named-range lookup для string row. Теперь оба Cells Invoke передают
полный вызов в Range.Cells handler с уже реализованными правилами Item.
IDL/IID/vtable/native getters не менялись. Calls без result pointer получают
temporary VARIANT с обязательным VariantClear; child HRESULT возвращается.

На обеих архитектурах new VBS49 checks PASS. C тест проверил точные HRESULT,
VT_EMPTY при ошибке, byref Double и BSTR ownership, ignored get results,
METHOD|PROPERTYGET, NULL DISPPARAMS, bad get/put argcount:0 failures. Новый
валидный get требует и error0, и non-Nothing object. Прежний legacy BSTR default
для Worksheet.Range сохранён; Cells использует numeric Item coercion.
Подробности: [Cells indexing](cells-indexing-results.md).

Final PE32/PE64: прежние Item/Cells ownership/parent/Worksheet.Range/Application.Range
probes PASS; полные7 legacy tests последовательно на обеих: counters158/0,
markers156/0. Outputs ~/tmp/claude/unioffice-excel-suite-0lyrmghp (32), -_ucysqo5
(64). MSI32/MSI64 пересобраны, architecture/Unicode metadata0; lifecycle не
повторялся (registration/IDL unchanged). Deployed DLL совпали с outputs побайтно.
Нативная Windows не проверялась; нового подтверждённого Wine defect нет.

## Range.Rows/Columns: границы и тип коллекции (2026-11-06)

Раньше get_Rows/get_Columns вызывали EntireRow/EntireColumn: терялись границы
исходного диапазона, Count/Item следовали cell rules. PE64 red probe получил
49 failures в71 выполненной проверке. Теперь отдельные views сохраняют UNO
rectangle и сильную ссылку на worksheet. Private collection_kind управляет
Count и Item/default selection: B2:D4.Rows(2) → B3:D3; Columns("AA") → AB2:AB4.
Второй Item index запрещён (E_INVALIDARG), положительные offsets могут выходить
за исходный rectangle. Cells создаёт отдельный cell view, не меняя source.
Group/Ungroup используют collection axis; прежняя width heuristic остаётся
только для обычных cell ranges. IDL/IIDs/public vtables не менялись; private
header forced rebuild пересобрал все зависимые объекты на обоих хостах.

Final PE32/PE64: new VBS86 PASS каждый, C COM0 failures (native slots, NULL,
Count/coordinates, byref Double, BSTR conversion/overflow, exact HRESULT/VT_EMPTY,
ignored result). Прежние Cells49/Item36/ownership/parent/Worksheet/Application
probes PASS. Все7 legacy tests последовательно оба: counters158/0, markers156/0,
включая Group/Ungroup. Outputs ~/tmp/claude/unioffice-excel-suite-x7x89q44 (32),
-bqcaa4i6 (64). MSI32/MSI64 rebuilt, metadata/Unicode0; lifecycle не повторялся
(registration/IDL unchanged). Fixture DLLs побайтно совпали с outputs.

Отчёт: [Rows/Columns](range-rows-columns-results.md). Это single-rectangle Range
selection/Count; Worksheet/Application Rows/Columns, enumeration, string slices,
nonpositive offsets, полный набор value/format/delete/insert semantics и native
Windows oracle не проверены. Нового подтверждённого Wine defect нет.

## Application.QueryInterface: опубликованная dual vtable (2026-11-06)

Новый native Rows/Columns тест выявил ошибку QueryInterface(IID__Application):
возвращался private Disp_Application с прямыми return values вместо опубликованных
HRESULT/out-parameter slots. PE32 возвращал S_OK/мусор без записи output, затем
тест упал. Это подтверждено сопоставлением QI, IDL, generated header и обеих vtable;
не Wine defect. Теперь IID__Application → APPEXCEL, канонический IUnknown остаётся
прежним DISPAPPEXCEL, остальные QI branches/IDL/registration не менялись.

Отдельный tests/application-interface.c проверяет guarded Creator output/HRESULT,
QI roundtrip/repeat, canonical IUnknown identity и CoCreateInstance dual IID.
Отчёт: [Application ABI](application-interface-results.md). Исходный failure log:
~/tmp/claude/unioffice-sheet-axes-32-native-before-qi.log.
Отдельный ABI probe и расширенный native COM test дали0 failures на PE32/PE64.
Логи -interface-final.log/-native-final.log с prefix unioffice-sheet-axes-{32,64}.

## Окружение и договорённости

- Репозиторий: `/srv/lav/Projects/obsoleted/UniOffice/unioffice`.
  Локально путь также отображается как `/home/lav/Projects/obsoleted/UniOffice/unioffice`.
- `ssh fedora`: Fedora видит тот же `/srv` через NFS. Правки общие;
  не создавать вторую копию и не запускать одновременно сборку одного каталога.
- Локально доступны `winegcc`, `wineg++`, `widl`, `xvfb-run`.
  Установлены `wine-etersoft` и `wine-etersoft-devel` версии
  `11.20.1-eter1.x86_64`; GCC сообщает `15.3.1`.
- На Fedora найден `/usr/bin/i686-w64-mingw32-gcc` и `/usr/bin/widl`.
  Наличие 64-битного MinGW не подтверждено.
- Временные файлы и тестовые префиксы создавать в `~/tmp/claude`.
  Для этой проверки создан `~/tmp/claude/unioffice-com-check/prefix`.
  Не использовать рабочий Wine-префикс для экспериментов с регистрацией COM.
- Задачи по доработке Wine заводить в `https://bugs.etersoft.ru/`.
  Ошибки UniOffice не выдавать за ошибки Wine. Перед созданием искать дубликаты.
- Согласовывать существенный план, выполнять обычные шаги самостоятельно.
  Сообщения коммитов — на английском; несвязанные изменения разделять.
  Для публикации использовать `gpush`, не `git push`.
- `AGENTS.md` создан ранее; пользователь запретил перезаписывать существующий файл.
  Актуальные сведения о новой сборке находятся в `README.md` и этой заметке.

## Проверенные команды сборки

```sh
# Локально: Excel и Word, ELF x86-64 и библиотеки типов.
make TOOLCHAIN=winegcc -j4

# Fedora: Excel и Word, PE32 i386 и библиотеки типов, без MSI.
ssh fedora 'cd /srv/lav/Projects/obsoleted/UniOffice/unioffice && make libraries TOOLCHAIN=mingw MINGW=i686-w64-mingw32 -j4'
```

Обе команды успешно выполнялись после добавления выбора компилятора.
Повторная сборка без изменений не выполняла лишних команд. После последующих
правок `LONG` и фабрики локальная сборка вновь прошла; повторная проверка
этих последних правок на Fedora ещё требуется.

- `TOOLCHAIN=mingw` оставлен по умолчанию; исторический `MINGW=i586-pc-mingw32`
  нужно переопределять на Fedora.
- `BITS`: 32 для MinGW, 64 для префикса `x86_64` и для winegcc по умолчанию.
- Результаты: `unioffice_{excel,word}/build/<toolchain>-<bits>/`.
  MinGW создаёт `.dll`, winegcc — `.dll.so`, оба — `.tlb`.
- `Makefile.mingw` выбирает инструменты; `Makefile.component` содержит общие
  правила генерации и компиляции. Имя первого файла оставлено историческим.
- Для Winelib нужен `-mno-cygwin` при компиляции и линковке: используются
  заголовки/функции MSVCRT, а не смесь с Unix libc.
- В Winelib `__declspec(dllexport)` игнорируется компилятором ELF.
  Добавлены `.spec` с четырьмя экспортами: `DllGetClassObject`,
  `DllCanUnloadNow`, `DllRegisterServer`, `DllUnregisterServer`.
  Символы проверены через `nm -D`; MinGW-экспорты — через `objdump -p`.
- `--add-stdcall-alias` передаётся MinGW через `-Wl,`, а не непосредственно GCC.
- Генерируемые заголовки включаются через `<...>`, чтобы старые файлы рядом
  с исходниками не перекрывали заголовки выбранной сборки.
- После изменения флагов/версии компилятора нужна очистка выбранного варианта:
  `make clean TOOLCHAIN=winegcc`. Флаги командной строки автоматически не хешируются.
- Неизвестный `TOOLCHAIN` отклоняется с понятной ошибкой.

## IDL и MSI

Современный `widl` отказался генерировать Excel TLB: `AppEvents` упоминался
в `coclass Application`, но нигде не определялся. Убрана декларация этого
не реализованного source interface; реализация событий не добавлялась.

На Fedora `widl` первоначально не находил `stdole2.tlb`. Найдены пути:

- `/usr/i686-w64-mingw32/sys-root/mingw/include/stdole2.tlb`;
- `/usr/lib64/wine-etersoft/x86_64-windows/stdole2.tlb`.

В Makefile добавлен поиск первого пути по `MINGW`, с передачей `-L` в `widl`.
Для других расположений можно задать `IDLFLAGS="-L /path/to/typelibs"`.

`make` для MinGW по-прежнему включает MSI; `make libraries` его пропускает.
MSI использует `wine98`, `winepath` и утилиты в `MSI/bin/`, упаковывает только
Excel. Путь к новой сборке передаётся через `EXCEL_BUILD`; пути CABARC
переводятся `winepath -w`. `make msi TOOLCHAIN=winegcc` запрещён: ELF-библиотека
не является Windows DLL для этого установщика. MSI после правок не проверялся.

## Ошибки UniOffice и внесённые исправления

### Размеры типов

На Unix x86-64 `long` имеет 64 бита, Windows `LONG` — 32. IDL `long` порождает
`LONG`; прежние C-реализации с `long *` не соответствовали COM-vtable и могли
записывать лишние байты. В исходниках и заголовках Excel заменены C-типы `long`
на `LONG`, включая связанные поля, локальные значения и прототипы.
IDL, строки, комментарии и исходные кодировки сохранены. Это не полный аудит
64-битной совместимости; остальные предупреждения требуют отдельного разбора.

Добавлен доступный всем реализациям прототип `AutoWrap` из `tools.h`.
В `debug.h` явно объявлен `extern BOOL write_log` и подключён `<wchar.h>`.

### Жизненный цикл COM

`OOFFICE_ClassFactory` — статический объект, но `Release` вызывал для него
`HeapFree`. Этот вызов убран. В `QueryInterface` и `CreateInstance` фабрики
теперь обнуляется выходной указатель перед отказом. `DllGetClassObject`
проверяет `ppv` на `NULL` до записи и возвращает `E_POINTER`.

Изменения пока проверены компиляцией. Тест не дошёл до вызовов фабрики из-за
проблемы загрузки Winelib DLL (см. ниже); не считать их runtime-проверенными.

При чтении также замечено: в `DllMain` при освобождении многих `ti_*`
используется `ti_excel->lpVtbl`, а не vtable соответствующего объекта.
Это ещё не исправлено и требует отдельной проверки.

## COM smoke test и фактические результаты

Добавлен `tests/com-smoke.c`. Тест принимает путь DLL и TLB, не регистрирует COM
и не требует установленного OpenOffice. Он проверяет экспорты, неизвестный CLSID,
нулевой выходной указатель, отказы QueryInterface/aggregation, циклы получения
и освобождения фабрики, LockServer/CanUnloadNow и чтение TLB без регистрации.

```sh
winegcc -m64 -mno-cygwin \
  -o unioffice_excel/build/winegcc-64/com-smoke.exe \
  tests/com-smoke.c -lole32 -loleaut32 -luuid
```

`com-smoke.exe` здесь shell-запускатель, реальный ELF — `com-smoke.exe.so`.
Запускать запускатель напрямую, а не передавать shell-файл как Windows EXE.
Обычная обёртка `/usr/bin/wine` без DISPLAY/WAYLAND_DISPLAY отказалась работать;
используется `xvfb-run -a`, `TMPDIR=~/tmp/claude`, отдельный `WINEPREFIX` и timeout.
Предупреждения Mesa/DRI3 в Xvfb сами по себе не означают провал COM-теста.

### ABI функции main: кандидат на задачу Wine

При `int main(int argc, char **argv)` тест получил `argc=7215392`, хотя
`GetCommandLineA()` показывал корректные два аргумента. После явного
`int __cdecl main(...)` обработка аргументов заработала и выполнение дошло
до `LoadLibraryA`. В выводе `winegcc -v` видно:
`__cdecl=__stdcall`, `__stdcall=__attribute__((ms_abi))` для `-m64`.

Это подтверждённое наблюдение на smoke-тесте, но пока не доказанная причина
в Wine: нужен минимальный пример, проверка ABI startup-кода и воспроизведение
с/без `__cdecl`. Дубликаты по слову `winegcc` искались в Etersoft:
найдена только закрытая задача 16741 «winegcc: Could not find g++», она о другом.
Новая задача пока не создана.

### Загрузка DLL: открытый вопрос

`LoadLibraryA` вернул ошибку 126 при абсолютном Windows-пути к имени `.dll`
(на диске есть `.dll.so`). Попытка с именем `unioffice_excel.dll` и
`WINEDLLPATH`, указывающим на каталог сборки, также дала 126.
Проверка абсолютного пути с суффиксом `.dll.so` завершилась по timeout
через 30 секунд (код 124), без итогового отчёта smoke-теста. Все запущенные
в этой сессии тестовые команды завершены. Причина пока не установлена:
не утверждать, что Wine вообще не поддерживает ELF DLL.

Для дальнейшего разбора найден исходник
[загрузчика upstream Wine](https://github.com/wine-mirror/wine/blob/master/dlls/ntdll/unix/loader.c).
Он содержит обработку SO DLL и отдельный `load_so_dll`; это ориентир для анализа,
а не доказательство поведения установленной сборки Etersoft.

## Ограничения существующих интеграционных тестов

`tests/Excel/all_test.vbs` запускает цепочку `test_<number>.vbs` асинхронно.
Нужно дождаться конца цепочки и читать `otchet.txt`; код завершения запускателя
не является результатом всех тестов. Запускать из `tests/Excel/` через
`cscript.exe /E:vbscript all_test.vbs` в настроенном окружении.

`_ApplicationConstructor` обращается к COM-серверу
`com.sun.star.ServiceManager` и создаёт `com.sun.star.frame.Desktop`.
Без совместимого зарегистрированного OpenOffice нельзя считать полный тест
Excel.Application доступным. Word `DllGetClassObject` пока возвращает
`E_NOTIMPL`: успешная сборка Word не означает готовую реализацию Word Automation.

## Что делать дальше

Семь исходных тестов с Windows OpenOffice 4.1.17 x86 проходят. В рамках
текущего набора регрессионных отказов не осталось. Возможные отдельные задачи:

1. MSI32/MSI64 и covered Excel PE32/PE64 сценарии проверены. Нативная Windows
   и другие версии Windows Installer потребуют отдельного окружения.
2. Borders.Weight реализован, включая mixed Null в Automation без поломки
   прежней native vtable. Точный порядок обхода Excel и прочие неподдержанные
   интерфейсы не покрыты; Word содержит заглушки. Отсутствие extended border API в Windows AOO4.1.17
   подтверждено; dashed/dotted styles дают E_NOTIMPL, включая aggregate/Value.
   Для поддержки пунктиров нужен отдельный проверенный backend с extended API.
   Worksheet.Range/Cells normalization и parent lifetime проверены и исправлены.
   Range.Item и single-rectangle Range.Rows/Columns selection/Count реализованы;
   Worksheet/Application Rows/Columns, multi-area, string slices, enumeration
   и прочие normalization sites остаются отдельными возможными работами.
3. При новых изменениях сохранять этапы отдельными коммитами, обновлять знания
   win-builder; winegcc на builder64, MinGW на Fedora.

Существовавшие до работы незакоммиченные `.msi`, сгенерированные файлы Word,
`IDL.segfault`, swap и временный заголовок Excel не удалялись. Накопленные изменения зафиксированы отдельными коммитами: PE-сборка/ABI
`2d4e460`, жизненный цикл COM `4d2e496`, MSI `1ef4a4d`, тесты/документация
`a8621dd`, выбор отдельных тестов `0abc02d`. Публикация не выполнялась. Учёт `work_time` при оформлении задач вести только
за соответствующий новый этап, не суммировать предыдущие этапы.
