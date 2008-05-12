#include <stdio.h>

/* TODO: print __FUNCTION__ */
#define DEBUG 1
#ifdef DEBUG
#define TRACE(args...) \
do { fprintf(stderr, args); } while(0)
#else
#define TRACE(n,...)
#endif

#ifdef DEBUG
#define WTRACE(args...) \
do { fwprintf(stderr, args); } while(0)
#else
#define WTRACE(n,...)
#endif

