# TODO: remove msvcrt using

GCC=i386-mingw32msvc-g++
LD=i386-mingw32msvc-ld
STRIP=i386-mingw32msvc-strip
DLLWRAP=i386-mingw32msvc-dlltool


LIBS=-ladvapi32 -lole32 -loleaut32 -luser32 -lkernel32 -lshell32 -luuid -lshlwapi -lmsvcrt
CFLAGS=-O2 -D_FORTIFY_SOURCE=2
OBJECTS = unioffice_word_i.o \
	main.o

.PHONY: all clean

all: unioffice_word.dll

.c.o:
	$(GCC) $(CFLAGS) -c $^ -o $@ 

unioffice_word_i.c:
	widl -I /usr/include/wine/windows -u -h -t unioffice_word.idl
unioffice_word.dll: $(OBJECTS)  
	$(LD) --add-stdcall-alias -shared -o $@ $^ $(LIBS)
#	$(DLLWRAP) --add-stdcall-alias --base-file main.o --output-def test.def -o $@ $^ $(LIBS)
	$(STRIP) --strip-all $@

clean:
	rm -f *.o *.dll *.h~ *.c~ *.idl~ unioffice_word.h unioffice_word_i.c *.tlb
