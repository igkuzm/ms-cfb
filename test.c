/**
 * File              : test.c
 * Author            : Igor V. Sementsov <ig.kuzm@gmail.com>
 * Date              : 05.11.2022
 * Last Modified Date: 07.11.2022
 * Last Modified By  : Igor V. Sementsov <ig.kuzm@gmail.com>
 */

#include "ole2.h"
#include "../libdoc2/property_set.h"

#include <iconv.h>
#include <stdint.h>
#include <stdio.h>

static char* unicode_decode_iconv(const char *s, size_t len, iconv_t ic) {
	char* outbuf = 0;

	if(s && len && ic)
	{
		size_t outlenleft = len;
		int outlen = len;
		size_t inlenleft = len;
		const char* src_ptr = s;
		char* out_ptr = 0;

		size_t st; 
		outbuf = malloc(outlen + 1);

		if(outbuf)
		{
			out_ptr = outbuf;
			while(inlenleft)
			{
				st = iconv(ic, (char **)&src_ptr, &inlenleft, (char **)&out_ptr,(size_t *) &outlenleft);
				if(st == (size_t)(-1))
				{
					if(errno == E2BIG)
					{
						size_t diff = out_ptr - outbuf;
						outlen += inlenleft;
						outlenleft += inlenleft;
						outbuf = realloc(outbuf, outlen + 1);
						if(!outbuf)
						{
							break;
						}
						out_ptr = outbuf + diff;
					}
					else
					{
						free(outbuf), outbuf = NULL;
						break;
					}
				}
			}
		}
		outlen -= outlenleft;

		if(outbuf)
		{
			outbuf[outlen] = 0;
		}
	}
	return outbuf;
}

int prop_cb(void * user_data, uint32_t propid, uint32_t dwType, uint32_t * value){
	char * str = NULL;
	
	if (dwType == 30){
		char buf[BUFSIZ];
		strncpy(buf, (char*)value + 4, BUFSIZ);
		//str = cprecode(buf, CPRECODE_CP1251, CPRECODE_UTF8);
		//str = buf;
		iconv_t ic = iconv_open("UTF-8", "CP1251");
		str = unicode_decode_iconv(buf, strlen(buf), ic);
	}
	else if (dwType == 2){
		char buf[BUFSIZ];
		sprintf(buf, "%d", *(uint16_t*)value);
		str = buf;
	}
	else if (dwType == 3){
		char buf[BUFSIZ];
		sprintf(buf, "%d", *(uint32_t*)value);
		str = buf;
	}

	printf("PROP id: %d, type: %d, value: %s\n", propid, dwType, str);
	return 0;
};

int callback(void * user_data, cbf_dir dir){
	printf("DIR AB: %x\n", dir._ab[0]);

	return 0;
}

int main(int argc, char *argv[])
{

	struct cbf cbf;
	int error = cbf_open(&cbf, "1.doc");
	if (error)
		printf("ERROR OPEN FILE: %x\n", error);
	
	printf("ROOT DIR: %s\n", cbf_dir_name(&cbf.root));

	FILE *si = cbf_dir_get_stream_by_name(&cbf, "\005SummaryInformation");


	/*cbf_dir dir;*/
	/*cbf_dir_by_sid(&cbf, 4, &dir, cbf_dir_callback);*/
	
	/*if (cbf_get_dir_by_sid(&cbf, &dir, 4))*/
		/*printf("ERRROR TO OPEN DIR\n");*/

	/*if (cbf_get_dir_by_name(&cbf, &dir, "\005SummaryInformation"))*/
		/*printf("ERRROR TO OPEN DIR\n");	*/
	
	/*printf("OPENED DIR: %s\n", cbf_dir_name(&dir));*/
	
	/*FILE * stream = ole2_dir_stream(dir);	*/

	property_set_get(si, NULL, prop_cb);

	return 0;
}

