/**
 * File              : log.h
 * Author            : Igor V. Sementsov <ig.kuzm@gmail.com>
 * Date              : 20.02.2023
 * Last Modified Date: 20.02.2023
 * Last Modified By  : Igor V. Sementsov <ig.kuzm@gmail.com>
 */

#ifndef LOG_H_
#define LOG_H_

#ifdef __cplusplus
extern "C"{
#endif

#include <stdio.h>
/*
 * Debugging
 */
#ifndef logfile
#define logfile stderr
#endif
#ifndef LOG
#define LOG(...) ({fprintf(logfile, __VA_ARGS__);})
#endif	

#ifdef __cplusplus
}
#endif

#endif //LOG_H_

// vim:ft=c	
