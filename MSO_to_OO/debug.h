#include <stdio.h>

#define DEBUG 1
#ifdef DEBUG
#define TRACE(args...) \
do { fprintf(stderr,"%s:%s:",__FILE__,__FUNCTION__);fprintf(stderr, args); } while(0)
#else
#define TRACE(n,...)
#endif

#ifdef DEBUG
#define WTRACE(args...) \
do { fprintf(stderr,"%s:%s:",__FILE__,__FUNCTION__);fwprintf(stderr, args); } while(0)
#else
#define WTRACE(n,...)
#endif

