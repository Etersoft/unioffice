
include ../Makefile.mingw

IDL=widl

# TODO: remove msvcrt using
LIBS=-ladvapi32 -lole32 -loleaut32 -luser32 -lkernel32 -lshell32 -luuid -lshlwapi -lmsvcrt

OBJECTS = unioffice_word_i.o \
	regsvr.o \
	main.o

.PHONY: all clean

all: unioffice_word.dll

unioffice_word_i.c:
	$(IDL) -I /usr/include/wine/windows -u -h -t unioffice_word.idl

unioffice_word.dll: $(OBJECTS)
	$(CXX) --add-stdcall-alias -shared -o $@ $^ $(LIBS)
	$(STRIP) --strip-all $@

clean:
	rm -f *.o *.dll *.h~ *.c~ *.cpp~ *.idl~ unioffice_word.h unioffice_word_i.c *.tlb
