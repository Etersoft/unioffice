# переименовать ver и name.idt

MSIDB=wine bin/MsiDb.exe
MAKECAB=wine bin/MAKECAB.EXE
CABARC=wine bin/CABARC.EXE n
PWD=`winepath -w $$(pwd)`

VAR_ver=`cat ver.idt`
VAR_name=`cat name.idt`

all: unioffice.msi
	$(MAKE) rmtmp

# Заполняем CAB-файлы содержимым
ooffice.cab: ./../MSO_to_OO/mso_to_oo.dll
	$(MAKECAB) ..\\MSO_to_OO\\mso_to_oo.dll $@

# Удаляем временные файлы
rmtmp:
	rm -f *.cab

# Удаляем все создаваемые файлы
clean: rmtmp
	rm -f unioffice.msi

# Извлекаем .idt из msi (если правили его на стороне)
extridt:
	$(MSIDB) -d $(PWD)\\Unioffice.msi -f $(PWD)\\idt -e "*"
	rm -f idt/_Validation.idt idt/_SummaryInformation.idt

# Формируем пакет MSI из шаблона, добавляем CAB'ы и .idt
unioffice.msi: bin/template.msi ooffice.cab
	cp bin/template.msi $@
	$(MSIDB) -d $@ -f $(PWD)\\idt -i "*"
	$(MSIDB) -d $@ -f $(PWD)\\idt\\Binary -i "*"
	$(MSIDB) -d $@ -a ooffice.cab
