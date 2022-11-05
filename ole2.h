/**
 * File              : ole2.h
 * Author            : Igor V. Sementsov <ig.kuzm@gmail.com>
 * Date              : 03.11.2022
 * Last Modified Date: 05.11.2022
 * Last Modified By  : Igor V. Sementsov <ig.kuzm@gmail.com>
 */

#ifndef OLE2_H_
#define OLE2_H_

#ifdef __cplusplus
extern "C"{
#endif

#include <stdint.h>
#include <stdio.h>
#include <stdbool.h>
#include <stdlib.h>
#include <string.h>
#include <errno.h>
#include <sys/types.h>

/*
 * A Compound File is made up of a number of virtual streams. These are collections of data that behave as a 
 * linear stream, although their on-disk format may be fragmented. Virtual streams can be user data, or they 
 * can be control structures used to maintain the file. Note that the file itself can also be 
 * considered a virtual stream.
 * All allocations of space within a Compound File are done in units called sectors. The size of a sector 
 * is definable at creation time of a Compound File, but for the purposes of this document will be 512 bytes. 
 * A virtual stream is made up of a sequence of sectors.
 * The Compound File uses several different types of sector: Fat, Directory, Minifat, DIF, and Storage. 
 * A separate type of 'sector' is a Header, the primary difference being that a Header is always 512 bytes long 
 * (regardless of the sector size of the rest of the file) and is always located at offset zero (0). 
 * With the exception of the header, sectors of any type can be placed anywhere within the file. The function 
 * of the various sector types is discussed below.
 * In the discussion below, the term SECT is used to describe the location of a sector within a virtual 
 * stream (in most cases this virtual stream is the file itself). Internally, a SECT is represented as a ULONG.
 */

/*
 * TIPES
 */

typedef uint32_t ULONG    ; //[4 bytes]
typedef uint16_t USHORT   ; //[2 bytes]
typedef uint16_t OFFSET   ; //[2 bytes]
typedef ULONG SECT        ; //[4 bytes]
typedef ULONG FSINDEX     ; //[4 bytes]
typedef USHORT FSOFFSET   ; //[2 bytes]
typedef ULONG DFSIGNATURE ; //[4 bytes]
typedef uint8_t BYTE      ; //[1 byte]
typedef uint16_t WORD     ; //[2 bytes]
typedef uint32_t DWORD    ; //[4 bytes]
typedef WORD DFPROPTYPE   ; //[2 bytes]
typedef ULONG SID         ; //[4 bytes]

#ifndef CLSID
typedef struct CLSID {
	uint32_t a;
	uint32_t b;
	uint32_t c;
	uint32_t d;
}CLSID;
#endif
typedef CLSID GUID        ;//[16 bytes]

typedef struct tagFILETIME {
	DWORD dwLowDateTime; 
	DWORD dwHighDateTime; 
} FILETIME, TIME_T        ;//[8 bytes]      


const SECT DIFSECT        = 0xFFFFFFFC; 
const SECT FATSECT        = 0xFFFFFFFD;
const SECT ENDOFCHAIN     = 0xFFFFFFFE;
const SECT FREESECT       = 0xFFFFFFFF;
 
/*
 * HEADER
 *
 * The Header contains vital information for the instantiation of a Compound File. 
 * Its total length is 512 bytes. There is exactly one Header in any Compound File, 
 * and it is always located beginning at offset zero in the file.
 *
 * Fat Sectors
 *
 * The Fat is the main allocator for space within a Compound File. Every sector in the file is 
 * represented within the Fat in some fashion, including those sectors that are unallocated 
 * (free). The Fat is a virtual stream made up of one or more Fat Sectors.
 * Fat sectors are arrays of SECTs that represent the allocation of space within the file. Each 
 * stream is represented in the Fat by a chain, in much the same fashion as a DOS file-allocation-
 * table (FAT). To elaborate, the set of Fat Sectors can be considered together to be a single 
 * array -- each cell in that array contains the SECT of the next sector in the chain, and this 
 * SECT can be used as an index into the Fat array to continue along the chain. Special values 
 * are reserved for chain terminators (ENDOFCHAIN = 0xFFFFFFFE), free sectors (FREESECT = 
 * 0xFFFFFFFF), and sectors that contain storage for Fat Sectors (FATSECT = 0xFFFFFFFD) or DIF 
 * Sectors (DIFSECT = 0xFFFFFFC), which are not chained in the same way as the others.
 * The locations of Fat Sectors are read from the DIF (Double-indirect Fat), which is described 
 * below. The Fat is represented in itself, but not by a chain – a special reserved SECT value 
 * (FATSECT = 0xFFFFFFFD) is used to mark sectors allocated to the Fat.
 * A SECT can be converted into a byte offset into the file by using the following formula: SECT 
 * << ssheader._uSectorShift + sizeof(ssheader). This implies that sector 0 of the file begins at 
 * byte offset 512, not at 0.
 *
 * MiniFat Sectors
 * Since space for streams is always allocated in sector-sized blocks, there can be considerable 
 * waste when storing objects much smaller than sectors (typically 512 bytes). As a solution to 
 * this problem, we introduced the concept of the MiniFat. The MiniFat is structurally equivalent 
 * to the Fat, but is used in a different way. The virtual sector size for objects represented in 
 * the Minifat is 1 << ssheader._uMiniSectorShift (typically 64 bytes) instead of 1 << ssheader.
 * _uSectorShift (typically 512 bytes). The storage for these objects comes from a virtual stream 
 * within the Multistream (called the Ministream).
 * The locations for MiniFat sectors are stored in a standard chain in the Fat, with the 
 * beginning of the chain stored in the header.
 * A Minifat sector number can be converted into a byte offset into the ministream by using the 
 * following formula: SECT << ssheader._uMiniSectorShift. (This formula is different from the 
 * formula used to convert a SECT into a byte offset in the file, since no header is stored in 
 * the Ministream).
 * The Ministream is chained within the Fat in exactly the same fashion as any normal stream. It 
 * is referenced by the first Directory Entry (SID 0).
 *
 * DIF Sectors
 * The Double-Indirect Fat is used to represent storage of the Fat. The DIF is also represented 
 * by an array of SECTs, and is chained by the terminating cell in each sector array. As an 
 * optimization, the first 109 Fat Sectors are represented within the header itself, so no DIF 
 * sectors will be found in a small (< 7 MB) Compound File.
 * The DIF represents the Fat in a different manner than the Fat represents a chain. A given 
 * index into the DIF will contain the SECT of the Fat Sector found at that offset in the Fat 
 * virtual stream. For instance, index 3 in the DIF would contain the SECT for Sector #3 of the 
 * Fat.
 * The storage for DIF Sectors is reserved in the Fat, but is not chained there (space for it is 
 * reserved by a special SECT value , DIFSECT=0xFFFFFFFC). The location of the first DIF sector 
 * is stored in the header.
 * A value of ENDOFCHAIN=0xFFFFFFFE is stored in the pointer to the next DIF sector of the last 
 * DIF sector.
 */ 

static const char wcbff_signature[8]     = {0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1};
static const char wcbff_signature_old[8] = {0x0e, 0x11, 0xfc, 0x0d, 0xd0, 0xcf, 0x11, 0xe0};
 
struct StructuredStorageHeader {  // [offset from start in bytes, length in bytes]
	BYTE _abSig[8];               // [000H,08] {0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1} for current version,
                                  //       was {0x0e, 0x11, 0xfc, 0x0d, 0xd0, 0xcf, 0x11, 0xe0} on old, beta 2 files (late ’92) 
								  //       which are also supported by the reference implementation
	CLSID _clid;                  // [008H,16] class id (set with WriteClassStg, retrieved with GetClassFile/ReadClassStg)
	USHORT _uMinorVersion;        // [018H,02] minor version of the format: 33 is written by reference implementation
	USHORT _uDllVersion;          // [01AH,02] major version of the dll/format: 3 is written by reference implementation
	USHORT _uByteOrder;           // [01CH,02] 0xFFFE: indicates Intel byte-ordering
	USHORT _uSectorShift;         // [01EH,02] size of sectors in power-of-two (typically 9, indicating 512-byte sectors)
	USHORT _uMiniSectorShift;     // [020H,02] size of mini-sectors in power-of-two (typically 6, indicating 64-byte mini-sectors)
	USHORT _usReserved;           // [022H,02] reserved, must be zero
	ULONG  _ulReserved1;          // [024H,04] reserved, must be zero
	ULONG  _ulReserved2;          // [028H,04] reserved, must be zero
	FSINDEX _csectFat;            // [02CH,04] number of SECTs in the FAT chain
	SECT	_sectDirStart;        // [030H,04] first SECT in the Directory chain
	DFSIGNATURE _signature;       // [034H,04] signature used for transactionin: must be zero. The reference implementation
								  // does not support transactioning
	ULONG _ulMiniSectorCutoff;    // [038H,04] maximum size for mini-streams: typically 4096 bytes
	SECT  _sectMiniFatStart;      // [03CH,04] first SECT in the mini-FAT chain
	FSINDEX _csectMiniFat;        // [040H,04] number of SECTs in the mini-FAT chain
	SECT  _sectDifStart;          // [044H,04] first SECT in the DIF chain
	FSINDEX _csectDif;            // [048H,04] number of SECTs in the DIF chain
	SECT _sectFat[109];           // [04CH,436] the SECTs of the first 109 FAT sectors
};
typedef struct StructuredStorageHeader wcbff_header;


/*
 * Directory Sectors
 *
 * The Directory is a structure used to contain per-stream information about the streams in 
 * a Compound File, as well as to maintain a tree-styled containment structure. It is a 
 * virtual stream made up of one or more Directory Sectors. The Directory is represented as 
 * a standard chain of sectors within the Fat. The first sector of the Directory chain 
 * (the Root Directory Entry)
 *
 * Each level of the containment hierarchy (i.e. each set of siblings) is represented as a 
 * red-black tree. The parent of this set of sibilings will have a pointer to the top of 
 * this tree. This red-black tree must maintain the following conditions in order for it 
 * to be valid:
 * 1. The root node must always be black. Since the root directory (see below) does not 
 * have siblings, it's color is irrelevant and may therefore be either red or black.
 * 2. No two consecutive nodes may both be red.
 * 3. The left child must always be less than the right child. This relationship is 
 * defined as:
 *		- A node with a shorter name is less than a node with a longer name (i.e. compare 
 *length of the name)
 *		- For nodes with the same length names, compare the two names.
 * The simplest implementation of the above invariants would be to mark every node as black, 
 * in which case the tree is simply a binary tree.
 * A Directory Sector is an array of Directory Entries, a structure represented in the diagram
 * below. Each user stream within a Compound File is represented by a single Directory Entry.
 * The Directory is considered as a large array of Directory Entries. It is useful to note 
 * that the Directory Entry for a stream remains at the same index in the Directory array for 
 * the life of the stream – thus, this index (called an SID) can be used to readily identify 
 * a given stream.
 * The directory entry is then padded out with zeros to make a total size of 128 bytes.
 * Directory entries are grouped into blocks of four to form Directory Sectors.
 *
 * Root Directory Entry
 *
 * The first sector of the Directory chain (also referred to as the first element of the
 * Directory array, or SID 0) is known as the Root Directory Entry and is reserved for two
 * purposes: First, it provides a root parent for all objects stationed at the root of the 
 * multi-stream. Second, its function is overloaded to store the size and starting sector for 
 * the Mini-stream.
 * The Root Directory Entry behaves as both a stream and a storage. All of the fields in the
 * Directory Entry are valid for the root. The Root Directory Entry’s Name field typically
 * contains the string “RootEntry” in Unicode, although some versions of structured storage
 * (particularly the preliminary reference implementation and the Macintosh version) store only
 * the first letter of this string, “R” in the name. This string is always ignored, since the 
 * Root Directory Entry is known by its position at SID 0 rather than by its name, and its name
 * is not otherwise used. New implementations should write “RootEntry” properly in the Root
 * Directory Entry for consistency and support manipulating files created with only the “R” name.
 *
 * Other Directory Entries
 *
 * Non-root directory entries are marked as either stream (STGTY_STREAM) or storage
 * (STGTY_STORAGE) elements. Storage elements have a _clsid, _time[], and _sidChild values;
 * stream elements may not. Stream elements have valid _sectStart and _ulSize members, whereas
 * these fields are set to zero for storage elements (except as noted above for the Root Directory Entry).
 * To determine the physical file location of actual stream data from a stream directory entry,
 * it is necessary to determine which FAT (normal or mini) the stream exists within. Streams 
 * whose _ulSize member is less than the _ulMiniSectorCutoff value for the file exist in the 
 * ministream, and so the _startSect is used as an index into the MiniFat (which starts at 
 * _sectMiniFatStart) to track the chain of mini-sectors through the mini-stream (which is, as 
 * noted earlier, the standard (non-mini) stream referred to by the Root Directory Entry’s 
 * _sectStart value). Streams whose _ulSize member is greater than the _ulMiniSectorCutoff value 
 * for the file exist as standard streams – their _sectStart value is used as an index into the 
 * standard FAT which describes the chain of full sectors containing their data).
 *
 * Storage Sectors
 * Storage sectors are simply collections of arbitrary bytes. They are the building blocks of user 
 * streams, and no restrictions are imposed on their contents. Storage sectors are represented as 
 * chains in the Fat, and each storage chain (stream) will have a single Directory Entry associated 
 * with it.
 */

typedef enum tagSTGTY { 
	STGTY_INVALID   = 0,
	STGTY_STORAGE   = 1,
	STGTY_STREAM    = 2,
	STGTY_LOCKBYTES = 3,
	STGTY_PROPERTY  = 4,
	STGTY_ROOT      = 5,
} STGTY;

typedef enum tagDECOLOR { 
	DE_RED          = 0,
	DE_BLACK        = 1,
} DECOLOR;


struct StructuredStorageDirectoryEntry {// [offset from start in bytes, length in bytes]                                               
	BYTE _ab[32*sizeof(WORD)];          // [000H,64] 64 bytes. The Element name in Unicode, 
										// padded with zeros to fill this byte array
	WORD _cb;                           // [040H,02] Length of the Element name in characters, not bytes
	BYTE _mse;                          // [042H,01] Type of object: value taken from the STGTY enumeration
	BYTE _bflags;                       // [043H,01] Value taken from DECOLOR enumeration.
	SID _sidLeftSib;                    // [044H,04] SID of the left-sibling of this entry in the directory tree
	SID _sidRightSib;                   // [048H,04] SID of the right-sibling of this entry in the directory tree
	SID _sidChild;                      // [04CH,04] SID of the child acting as the root of all the children of this element (if _mse=STGTY_STORAGE)
	GUID _clsId;                        // [050H,16] CLSID of this storage (if _mse=STGTY_STORAGE)
	DWORD _dwUserFlags;                 // [060H,04] User flags of this storage (if _mse=STGTY_STORAGE)
	TIME_T _time[2];                    // [064H,16] Create/Modify time-stamps (if _mse=STGTY_STORAGE) 
	SECT _sectStart;                    // [074H,04] starting SECT of the stream (if _mse=STGTY_STREAM) 
	ULONG _ulSize;                      // [078H,04] size of stream in bytes (if _mse=STGTY_STREAM)
    DFPROPTYPE _dptPropType;            // [07CH,02] Reserved for future use. Must be zero.                                           
};

/*
 * OLE2 Structures
 */

struct ole2 {
	FILE * fp;
	struct StructuredStorageHeader header;
	SECT startOfDir;
	unsigned sect_size;
	unsigned mini_sect_size;
	SECT startOfMiniFat;
	int sizeOfDir;
}; 
typedef struct ole2 * ole2_t;

struct ole2_dir {
	struct StructuredStorageDirectoryEntry dir;
	ole2_t ole2;
};
typedef struct ole2_dir * ole2_dir_t;

typedef struct ole2_dir_list {
	ole2_dir_t dir;
	struct ole2_dir_list * next;
} ole2_dir_list;

ole2_dir_list * ole2_dir_list_add(ole2_dir_list ** list, ole2_dir_t dir){
	//create list if needed
	if(*list == NULL){
		*list = malloc(sizeof(ole2_dir_list));
		if(*list == NULL)
			return NULL;
		list[0]->next = NULL;
	}

	//get last item in list
	ole2_dir_list * ptr = *list;
	while(ptr->next)
		ptr=ptr->next;

	//create new item
	ole2_dir_list * node  = malloc(sizeof(ole2_dir_list));
	if(!node)
		return NULL;
	node->next = NULL;
	ptr->next = node;
	ptr->dir = dir;

	return *list;
}

void ole2_dir_list_free(ole2_dir_list * list){
	while(list){
		ole2_dir_list * ptr=list->next;	
		if (list->dir)
			free(list->dir);
		free(list);
		list = ptr;
	}
}

#define ole2_list_for_each(item, list)\
	ole2_dir_list * ___ptr = list; \
	ole2_dir_t item; \
	for (item = ___ptr->dir; ___ptr->next; ___ptr=___ptr->next, item = ___ptr->dir)

/*
 * IMP
 */

//check if file is Windows Compound File
bool _check_ole2_signature(ole2_t ole2){
	int i;
	bool res = true;
	char * ptr = (char *)(ole2->header._abSig);

	for (i=0; i<8; i++){
		if ((*ptr != wcbff_signature[i] && *ptr != wcbff_signature_old[i])){
			res = false;
			break;
		}
		ptr++;
	}
	
	return res;
};

int _read_ole2_header(FILE * fp, struct StructuredStorageHeader * header){
	int i;

	//read 512 bites
	size_t size = fread(header, 1, 512, fp);
	if (size == 512)
		return 0;

	//free header
	free(header);
	return -1;
}

//return len of utf8 string
size_t _utf16_to_utf8(WORD * utf16, int len, char * utf8){
	int i, k = 0;
	for (i = 0; i < len/2; ++i) {
		WORD wc = utf16[i];
		if (wc > 0x80){ //2-byte
			//get first byte - first 5 bit 00000111 11000000
			//and mask with 11000000 
			utf8[k++] = ((wc & 0x7C0)>> 6) | 0xC0;

			//get last - 00000000 00111111 
			//and mask with 10000000 
			utf8[k++] = ( wc & 0x3F)       | 0x80;
		}
		else //ANSY
			utf8[k++] = wc;
	}
	//null-terminate
	utf8[k] = 0;	

	return k;
}

//resturn len of utf16 string
size_t _utf8_to_utf16(const char * utf8, int len, WORD * utf16){
	int i;
	char *ptr = (char *)utf8;
	while (*ptr){ //iterate chars
		
		uint8_t utf8_char[4];

		//get utf32
		uint16_t utf16_char;
		if ((*ptr & 0b11110000) == 240) {
			//take last 3 bit from first char
			uint16_t byte0 = (*ptr++ & 0b00000111) << 18;	
			
			//take last 6 bit from second char
			uint16_t byte1 = (*ptr++ & 0b00111111) << 12;	
			
			//take last 6 bit from third char
			uint16_t byte2 = (*ptr++ & 0b00111111) << 6;	
			
			//take last 6 bit from forth char
			uint16_t byte3 = *ptr++ & 0b00111111;	
			
			utf16_char = (byte0 | byte1 | byte2 | byte3);					
		} 
		else if ((*ptr & 0b11100000) == 224) {
			//take last 4 bit from first char
			uint16_t byte0 = (*ptr++ & 0b00001111) << 12;	
			
			//take last 6 bit from second char
			uint16_t byte1 = (*ptr++ & 0b00111111) << 6;	
			
			//take last 6 bit from third char
			uint16_t byte2 = *ptr++ & 0b00111111;	

			utf16_char = (byte0 | byte1 | byte2);
		} 
		else if ((*ptr & 0b11000000) == 192){
			//take last 5 bit from first char
			uint16_t byte0 = (*ptr++ & 0b00011111) << 6;	
			
			//take last 6 bit from second char
			uint16_t byte1 = *ptr++ & 0b00111111;	

			utf16_char = (byte0 | byte1);
		}
		else {
			utf16_char = *ptr++;
		} 				
					
		utf16[i++] = utf16_char;
	}
	return i;
}

FILE * _ole2_get_stream(ole2_t ole2, ULONG start, ULONG size){

	FILE * stream = tmpfile();

	//seek to start
	fseek(ole2->fp, start, SEEK_SET);

	ULONG i = 0;
	//copy data
	for(i=0; i<size; i+=sizeof(SECT)){
		if (i==size){
			fputc(EOF, stream);
			break;
		}

		SECT ch;
		fread(&ch, 1, sizeof(SECT), ole2->fp);

		if ((int)ch == EOF){
			fputc(EOF, stream);
			break;
		}

		///* BYTE ENDIAN check:  <04-11-22, yourname> */
		///* FAT chain append check:  <04-11-22, yourname> */
		/*if (ch == ENDOFCHAIN)*/
			/*printf("ENDOFCHAIN\n");*/

		fwrite(&ch, 1,  sizeof(SECT), stream);
	}

	fseek(stream, 0, SEEK_SET);
	return stream;
}


ole2_t ole2_init(FILE * fp){

	//allocate
	ole2_t ole2 = malloc(sizeof(ole2_t));
	if (!ole2)
		return NULL;
	
	//init with zeros
	memset(ole2, 0, sizeof(struct ole2));

	ole2->fp = fp;

	//init header
	if (_read_ole2_header(fp, &ole2->header)){
		free(ole2);
		return NULL;
	}
	
	if (!_check_ole2_signature(ole2)){ //not WCBF file 
		free(ole2);
		return NULL;
	}

	ole2->sizeOfDir = sizeof(struct StructuredStorageDirectoryEntry);
	
	ole2->sect_size = 1<<ole2->header._uSectorShift;
	ole2->mini_sect_size = 1<<ole2->header._uMiniSectorShift;
	
	ole2->startOfDir = ole2->header._sectDirStart * ole2->sect_size;
	ole2->startOfMiniFat = ole2->header._sectMiniFatStart * ole2->mini_sect_size;
	
	return ole2;
}


ole2_dir_t _ole2_dir_init(ole2_t ole2, SID sid){
	int i;
	
	//allocate dir
	ole2_dir_t dir = malloc(sizeof(struct ole2_dir));
	if (!dir)
		return NULL;	

	//init with zeros
	memset(dir, 0, sizeof(struct ole2_dir));

	//goto dir data
	fseek(ole2->fp, 
			512 + 
			ole2->startOfDir + 
			sid*(sizeof(struct StructuredStorageDirectoryEntry)), 
	SEEK_SET);
	
	//copy dir data
	int len = fread(&dir->dir, 1, ole2->sizeOfDir, ole2->fp);
	if (len != ole2->sizeOfDir){
		free(dir);
		return NULL;
	}

	dir->ole2 = ole2;

	return dir;
}

ole2_dir_t ole2_dir_child(ole2_dir_t dir){
	if (!dir->dir._sidChild) //dir has no childs
		return NULL;
	
	ole2_dir_t child = _ole2_dir_init(dir->ole2, dir->dir._sidChild);
	return child;	
}

ole2_dir_t wcbff_dir_left(ole2_dir_t dir){
	if (!dir->dir._sidLeftSib) //dir has no left
		return NULL;
	
	ole2_dir_t left = _ole2_dir_init(dir->ole2, dir->dir._sidLeftSib);
	return left;	
}

ole2_dir_t wcbff_dir_right(ole2_dir_t dir){
	if (!dir->dir._sidRightSib) //dir has no right
		return NULL;
	
	ole2_dir_t right = _ole2_dir_init(dir->ole2, dir->dir._sidRightSib);
	return right;	
}

char * ole2_dir_name(ole2_dir_t dir){
	char * name = malloc(2*dir->dir._cb);
	if (!name)
		return NULL;

	WORD * ab = (WORD*)&(dir->dir._ab);

	if (!_utf16_to_utf8(ab, dir->dir._cb, name))
		return NULL;

	return name;
}

ole2_dir_t _ole2_dir_find(ole2_dir_t dir, const char * name){
	if(!dir)
		return NULL;
	//check name
	char * dirname = ole2_dir_name(dir);
	int res;
	size_t name_len = strlen(name);
	size_t dirname_len = strlen(dirname);
	if (name_len < dirname_len)
		res = -1;
	else if (name_len > dirname_len)
		res = 1;
	else 
		res = strcmp(name, dirname);
	free(dirname);
	if (res == 0)
		return dir;
	if (res < 0){
		//check left
		if (dir->dir._sidLeftSib > 0){
			ole2_dir_t new_dir =  _ole2_dir_init(dir->ole2, dir->dir._sidLeftSib);
			free(dir);
			return _ole2_dir_find(new_dir, name);
		}
	} else {
		//check right
		if (dir->dir._sidRightSib > 0){
			ole2_dir_t new_dir =  _ole2_dir_init(dir->ole2, dir->dir._sidRightSib);
			free(dir);
			return _ole2_dir_find(new_dir, name);
		}	
	}
	return NULL;
}


#define wcbff_get_dir_id(wcbff, sid) _wcbff_dir_init(wcbff, sid)

ole2_dir_t ole2_get_dir(ole2_t ole2, const char * name){
	ole2_dir_t root = _ole2_dir_init(ole2, 0);
	ole2_dir_t find = _ole2_dir_find(ole2_dir_child(root), name);
	
	//free root
	free(root);
	return find;
}

#define ole2_dir(ole2, arg)\
	_Generic((arg), \
			char*: ole2_get_dir, \
			int:   _ole2_dir_init \
	)((ole2), (arg))	

FILE * ole2_dir_stream(ole2_dir_t dir) {
	ULONG size = dir->dir._ulSize;
	
	ULONG start;

	//check FAT or miniFAT
	if (size < dir->ole2->mini_sect_size){
		//use miniFAT
		start = 512 + 
				dir->ole2->startOfMiniFat +
				dir->dir._sectStart * dir->ole2->mini_sect_size;
	} else {
		//use FAT
		start = 512 + 
				dir->dir._sectStart * dir->ole2->sect_size;
	}

	return _ole2_get_stream(dir->ole2, start, size); 
};

ole2_t ole2_open(const char * filename){
	FILE * newfile;
	FILE * fp = fopen(filename, "r");
	if (!fp)
		return NULL;

	if (fseek(fp,0,SEEK_SET) == -1) {
		if ( errno == ESPIPE ) {
			//We got non-seekable file, create temp file
			if((newfile=tmpfile()) == NULL) {
				return NULL;
			}
			int ch;
			while ((ch = fgetc(fp)) != EOF)
				fputc(ch, newfile);
			
			fclose(fp);
			fseek(newfile,0,SEEK_SET);
		} else {
			return NULL;
		}
	} else {
		newfile=fp;
	}	
	
	return ole2_init(newfile); 
};

ole2_dir_list * ole2_dirs(ole2_t ole2){
	int i=0;
	ole2_dir_list * list = NULL;
	
	ole2_dir_t dir = _ole2_dir_init(ole2, i++);
	while(dir && dir->dir._ab[0]){
		ole2_dir_list_add(&list, dir);	
		dir = _ole2_dir_init(ole2, i++);
	}

	return list;
}

#ifdef __cplusplus
}
#endif

#endif //OLE2_H_
