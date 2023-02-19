/**
 * File              : byteorder.h
 * Author            : Igor V. Sementsov <ig.kuzm@gmail.com>
 * Date              : 20.02.2023
 * Last Modified Date: 20.02.2023
 * Last Modified By  : Igor V. Sementsov <ig.kuzm@gmail.com>
 */

#ifndef BYTEORDER_H_
#define BYTEORDER_H_

#ifdef __cplusplus
extern "C"{
#endif

#include <stdint.h>

//switch bite order
//64 bit
static uint64_t bo_64_sw (uint64_t i)
{
    unsigned char c1, c2, c3, c4, c5, c6, c7, c8;

	c1 = i & 255;
	c2 = (i >> 8) & 255;
	c3 = (i >> 16) & 255;
	c4 = (i >> 24) & 255;
	c5 = (i >> 32) & 255;
	c6 = (i >> 40) & 255;
	c7 = (i >> 48) & 255;
	c8 = (i >> 56) & 255;

	uint64_t k = ((uint64_t)c1 << 56) 
			 + ((uint64_t)c2 << 48) 
			 + ((uint64_t)c3 << 40) 
			 + ((uint64_t)c4 << 32) 
			 + ((uint64_t)c5 << 24) 
			 + ((uint64_t)c6 << 16) 
			 + ((uint64_t)c7 << 8) 
			 + c8;
	return k;
}

//32 bit
static uint32_t bo_32_sw (uint32_t i)
{
    unsigned char c1, c2, c3, c4;

	c1 = i & 255;
	c2 = (i >> 8) & 255;
	c3 = (i >> 16) & 255;
	c4 = (i >> 24) & 255;

	u_int32_t k = ((uint32_t)c1 << 24) 
		   + ((uint32_t)c2 << 16) 
			 + ((uint32_t)c3 << 8) 
			 + c4;
	return k;
}

//16 bit
static uint16_t bo_16_sw (uint16_t i)
{
    unsigned char c1, c2;
    
	c1 = i & 255;
	c2 = (i >> 8) & 255;

	uint16_t k = (c1 << 8) + c2;
	return k;
}


#ifdef __cplusplus
}
#endif

#endif //BYTEORDER_H_

// vim:ft=c	
