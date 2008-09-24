# TODO: remove msvcrt using

GCC=i386-mingw32msvc-gcc
LD=i386-mingw32msvc-ld
STRIP=i386-mingw32msvc-strip

LIBS=-ladvapi32 -lole32 -loleaut32 -luser32 -lkernel32 -lshell32 -luuid -lshlwapi -lmsvcrt
CFLAGS=-O2 -D_FORTIFY_SOURCE=2
OBJECTS = mso_to_oo_i.o \
	tools.o \
	main.o \
	factory.o \
	regsvr.o \
	i_border.o \
	i_borders.o \
	i_interior.o \
	i_font.o  \
	special_functions.o \
	i_workbooks.o \
	i_names.o \
	i_workbook.o \
	i_range.o \
	i_pagesetup.o\
	i_shape.o \
	i_shapes.o \
	i_sheets.o \
	i_outline.o\
	i_worksheet.o \
	application_excel.o

.PHONY: all clean

all: mso_to_oo.dll

.c.o:
	$(GCC) $(CFLAGS) -c $^ -o $@ 

mso_to_oo_i.c:
	widl -I /usr/include/wine/windows -u -h -t mso_to_oo.idl
mso_to_oo.dll: $(OBJECTS)  
	$(LD) --add-stdcall-alias -shared -o $@ $^ $(LIBS)
	$(STRIP) --strip-all $@

clean:
	rm -f *.o *.dll *.h~ *.c~ *.idl~ mso_to_oo.h mso_to_oo_i.c *.tlb
