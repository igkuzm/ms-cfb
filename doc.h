/**
 * File              : doc.h
 * Author            : Igor V. Sementsov <ig.kuzm@gmail.com>
 * Date              : 04.11.2022
 * Last Modified Date: 21.02.2023
 * Last Modified By  : Igor V. Sementsov <ig.kuzm@gmail.com>
 */

#ifndef DOC_H_
#define DOC_H_

#include "byteorder.h"
#include <stdbool.h>
#ifdef __cplusplus
extern "C"{
#endif

#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>

#include "cfb.h"

/*
 * [MS-DOC]: Word (.doc) Binary File Format
 * Specifies the Word (.doc) Binary File Format, which is the binary file format used by Microsoft 
 * Word 97, Microsoft Word 2000, Microsoft Word 2002, and Microsoft Office Word 2003.
 *
 * Characters 
 * The fundamental unit of a Word binary file is a character. This includes visual characters such as 
 * letters, numbers, and punctuation. It also includes formatting characters such as paragraph marks, 
 * end of cell marks, line breaks, or section breaks. Finally, it includes anchor characters such as 
 * footnote reference characters, picture anchors, and comment anchors. 
 * Characters are indexed by their zero-based Character Position, or CP.  
 * 
 * PLCs 
 * Many features of the Word Binary File Format pertain to a range of CPs. For example, 
 * a bookmark is a range of CPs that is named by the document author. As another 
 * example, a field is made up of three control characters with ranges of arbitrary document content 
 * between them. 
 * The Word Binary File Format uses a PLC structure to specify these and other kinds of 
 * ranges of CPs. A PLC is simply a mapping from CPs to other, arbitrary data. 
 *
 * Formatting 
 * The formatting of characters, paragraphs, sections, tables, and pictures is specified as a set of 
 * differences in formatting from the default formatting for these objects. Modifications to individual 
 * properties are expressed using a Prl. A Prl is a Single Property Modifier, or Sprm, and an operand 
 * that specifies the new value for the property. Each property has (at least) one unique Sprm that 
 * modifies 
 * it. For example, sprmCFBold modifies the bold formatting of text, and sprmPDxaLeft modifies the 
 * logical left indent of a paragraph. 
 * The final set of properties for text, paragraphs, and tables comes from a hierarchy of styles and 
 * from Prl elements applied directly (for example, by the user selecting some text and clicking the 
 * Bold button in the user interface). Styles allow complex sets of properties to be specified in a 
 * compact way. 
 * They also allow the user to change the appearance of a document without visiting every place in the 
 * document where a change is necessary. The style sheet for a document is specified by a STSH. 
 *
 * Tables 
 * A table consists of a set of paragraphs that has a particular set of properties applied. There are 
 * special characters that denote the ends of table cells and the ends of table rows, but there are no 
 * characters to denote the beginning of a table cell or the end of the table as a whole. Tables can be 
 * nested inside other tables.
 * 
 * Pictures 
 * Pictures in the Word Binary File format can be either inline or floating. An inline picture is 
 * represented by a character whose Unicode value is 0x0001 and has sprmCFSpec applied with a value of 
 * 1 and sprmCPicLocation applied to specify the location of the picture data. A floating picture is 
 * represented by an anchor character with a Unicode value of 0x0008 with sprmCFSpec applied with a 
 * value of 1. In addition, floating pictures are referenced by a PlcfSpa structure which contains 
 * additional data about the picture. A floating picture can appear anywhere on the same page as its 
 * anchor. The document author can choose to have the floating picture rearrange the text in various 
 * ways or to leave the text as is.
 *
 * The FIB 
 * The main stream of the Word Binary File Format begins with a File Information Block, or FIB. The FIB 
 * specifies the locations of all other data in the file. The locations are specified by a pair of 
 * integers, the first of which specifies the location and the second of which specifies the size. 
 * These integers appear in substructures of the FIB such as the FibRgFcLcb97. The location names are 
 * prefixed with fc; the size names are prefixed with lcb.
 *
 * Byte Ordering 
 * Some computer architectures number bytes in a binary word from left to right, which is referred to 
 * as big-endian. The bit diagram for this documentation is big-endian. Other architectures number the 
 * bytes in a binary word from right to left, which is referred to as little-endian. The underlying 
 * file format enumerations, objects, and records are little-endian
 */

/*
 * Error codes
 */
enum {
	DOC_NO_ERR,     //no error
	DOC_CB_STOP,    //stopped by callback
	DOC_ERR_FILE,   //error to read file
	DOC_ERR_HEADER, //error to read header
	DOC_ERR_ALLOC,  //memory allocation error
};


/*
 * The File Information Block.
 *
 * The Fib structure contains information about the document and specifies the file pointers to various 
 * portions that make up the document.
 * The Fib is a variable length structure. With the exception of the base portion which is fixed in 
 * size, every section is preceded with a count field that specifies the size of the next section.
 *
 * base (32 bytes): The FibBase. 
 * 
 * csw (2 bytes): An unsigned integer that specifies the count of 16-bit values corresponding to fibRgW 
 * that follow. MUST be 0x000E. 
 * 
 * fibRgW (28 bytes): The FibRgW97. 
 * 
 * cslw (2 bytes): An unsigned integer that specifies the count of 32-bit values corresponding
 * to fibRgLw that follow. MUST be 0x0016. 
 *
 * fibRgLw (88 bytes): The FibRgLw97. 
 *
 * cbRgFcLcb (2 bytes): An unsigned integer that specifies the count of 64-bit values corresponding to 
 * fibRgFcLcbBlob that follow. This MUST be one of the following values, depending on the value of 
 * nFib.
 */
struct nFib2cbRgFcLcb {
	uint16_t nFib;
	uint16_t cbRgFcLcb;
};

const struct nFib2cbRgFcLcb nFib2cbRgFcLcbTable[] = {
	{0x00C1, 0x005D}, 
	{0x00D9, 0x006C}, 
	{0x0101, 0x0088}, 
	{0x010C, 0x00A4}, 
	{0x0112, 0x00B7}, 
};

static int nFib2cbRgFcLcb_compare(const void *key, const void *value) {
    const struct nFib2cbRgFcLcb *cp1 = key;
    const struct nFib2cbRgFcLcb *cp2 = value;
    return cp1->nFib - cp2->nFib;
}

uint16_t cbRgFcLcb_get(uint16_t nFib){
    struct nFib2cbRgFcLcb *result = bsearch(&nFib, nFib2cbRgFcLcbTable,
            sizeof(nFib2cbRgFcLcbTable)/sizeof(nFib2cbRgFcLcbTable[0]),
            sizeof(nFib2cbRgFcLcbTable[0]), nFib2cbRgFcLcb_compare);
	if (result)
		return result->cbRgFcLcb;
	return 0;
}

/*
 * fibRgFcLcbBlob (variable): The FibRgFcLcb. 
 *
 * cswNew (2 bytes): An unsigned integer that specifies the count of 16-bit values corresponding to 
 * fibRgCswNew that follow. This MUST be one of the following values, depending on the value of nFib. 
 */

/*
 * fibRgCswNew (variable): If cswNew is nonzero, this is fibRgCswNew. Otherwise, it is not present
 * in the file. 
 */
struct nFib2cswNew {
	uint16_t nFib;
	uint16_t cswNew;
};

const struct nFib2cswNew nFib2cswNewTable[] = {
	{0x00C1, 0     }, 
	{0x00D9, 0x0002}, 
	{0x0101, 0x0002}, 
	{0x010C, 0x0002}, 
	{0x0112, 0x0005}, 
};

static int nFib2cswNew_compare(const void *key, const void *value) {
    const struct nFib2cswNew *cp1 = key;
    const struct nFib2cswNew *cp2 = value;
    return cp1->nFib - cp2->nFib;
}

uint16_t cswNew_get(uint16_t nFib){
    struct nFib2cswNew *result = bsearch(&nFib, nFib2cswNewTable,
            sizeof(nFib2cswNewTable)/sizeof(nFib2cswNewTable[0]),
            sizeof(nFib2cswNewTable[0]), nFib2cswNew_compare);
	if (result)
		return result->cswNew;
	return 0;
}

/*
 * FibBase 
 * The FibBase structure is the fixed-size portion of the Fib. 
 */
typedef struct FibBase 
{ 
    // Header 
    uint16_t   wIdent; //(2 bytes): An unsigned integer that specifies that this is a Word Binary File. 
					   //This value MUST be 0xA5EC. 
	uint16_t   nFib;   //(2 bytes): An unsigned integer that specifies the version number of the file 
					   //format used. Superseded by FibRgCswNew.nFibNew if it is present. 
					   //This value SHOULD be 0x00C1.
					   //A special empty document is installed with Word 97, Word 2000, Word 2002, 
					   //and Office Word 2003 to allow "Create New Word Document" from the operating 
					   //system. This document has an nFib of 0x00C0. In addition the BiDi build of 
					   //Word 97 differentiates its documents by saving 0x00C2 as the nFib. In both 
					   //cases treat them as if they were 0x00C1.
	uint16_t  unused;  //(2 bytes)
	uint16_t  lid;	   //(2 bytes): A LID that specifies the install language of the application that 
					   //is producing the document. If nFib is 0x00D9 or greater, then any East Asian 
					   //install lid or any install lid with a base language of Spanish, German or 
					   //French MUST be recorded as 0x0409. If the nFib is 0x0101 or greater, then any 
					   //install lid with a base language of Vietnamese, Thai, or Hindi MUST be 
					   //recorded as 0x0409.
	uint16_t  pnNext;  //(2 bytes): An unsigned integer that specifies the offset in the WordDocument 
					   //stream of the FIB for the document which contains all the AutoText items. If 
					   //this value is 0, there are no AutoText items attached. Otherwise the FIB is 
					   //found at file location pnNext×512. If fGlsy is 1 or fDot is 0, this value MUST 
					   //be 0. If pnNext is not 0, each FIB MUST share the same values for FibRgFcLcb97.
					   //fcPlcfBteChpx, FibRgFcLcb97.lcbPlcfBteChpx, FibRgFcLcb97.fcPlcfBtePapx, 
					   //FibRgFcLcb97.lcbPlcfBtePapx, and FibRgLw97.cbMac
	uint16_t ABCDEFGHIJKLM;
					   //A - fDot (1 bit): Specifies whether this is a document template.
					   //B - fGlsy (1 bit): Specifies whether this is a document that contains only 
					   //    AutoText items (see FibRgFcLcb97.fcSttbfGlsy, FibRgFcLcb97.fcPlcfGlsy and 
					   //    FibRgFcLcb97.fcSttbGlsyStyle).
					   //C - fComplex (1 bit): Specifies that the last save operation that was 
					   //    performed on this document was an incremental save operation.
					   //D - fHasPic (1 bit): When set to 0, there SHOULD be no pictures in the 
					   //    document. Picture watermarks could be present in the document even if 
					   //    fHasPic is 0.
					   //E - cQuickSaves (4 bits): An unsigned integer. If nFib is less than 0x00D9, 
					   //	 then cQuickSaves specifies the number of consecutive times this document 
					   //	 was incrementally saved. If nFib is 0x00D9 or greater, then cQuickSaves 
					   //	 MUST be 0xF
					   //F - fEncrypted (1 bit): Specifies whether the document is encrypted or 
					   //	 obfuscated as specified in Encryption and Obfuscation.
					   //G - fWhichTblStm (1 bit): Specifies the Table stream to which the FIB 
					   //	 refers. When this value is set to 1, use 1Table; when this value is set to 
					   //	 0, use 0Table.
					   //H - fReadOnlyRecommended (1 bit): Specifies whether the document author 
					   //	 recommended that the document be opened in read-only mode.
					   //I - fWriteReservation (1 bit): Specifies whether the document has a write-
					   //	 reservation password.
					   //J - fExtChar (1 bit): This value MUST be 1.
					   //K - fLoadOverride (1 bit): Specifies whether to override the language 
					   //	 information and font that are specified in the paragraph style at istd 0 
					   //	 (the normal style) with the defaults that are appropriate for the 
					   //	 installation language of the application.
					   //L - fFarEast (1 bit): Specifies whether the installation language of the 
					   //	 application that created the document was an East Asian language
					   //M - fObfuscated (1 bit): If fEncrypted is 1, this bit specifies whether the 
					   //	 document is obfuscated by using XOR obfuscation; otherwise, this bit MUST 
					   //	 be ignored	 
	uint16_t  nFibBack;//(2 bytes): This value SHOULD be 0x00BF. This value MUST be 0x00BF or 
					   //0x00C1. The nFibBack field is treated as if it is set to 0x00BF when a locale-
					   //specific version of Word 97 sets it to 0x00C1. 
	uint32_t  lKey;    //(4 bytes): If fEncrypted is 1 and fObfuscated is 1, this value specifies the 
					   //XOR obfuscation password verifier. If fEncrypted is 1 and fObfuscated is 0, 
					   //this value specifies the size of the EncryptionHeader that is stored at the 
					   //beginning of the Table stream as described in Encryption and Obfuscation. 
					   //Otherwise, this value MUST be 0. 
	uint8_t  envr;     //(1 byte): This value MUST be 0, and MUST be ignored. 
	uint8_t NOPQRS;
					   //N - fMac (1 bit): This value MUST be 0, and MUST be ignored.
					   //O - fEmptySpecial (1 bit): This value SHOULD be 0 and SHOULD be ignored.
					   //	 Word 97, Word 2000, Word 2002, and Office Word 2003 install a minimal .doc 
					   //	 file for use with the New- Microsoft Word Document of the shell. This 
					   //	 minimal .doc file has fEmptySpecial set to 1.
					   //	 Word uses this flag to identify a document that was created by using the 
					   //	 New – Microsoft Word Document of the operating system shell
					   //P - fLoadOverridePage (1 bit):  Specifies whether to override the section 
					   //	 properties for page size, orientation, and margins with the defaults that 
					   //	 are appropriate for the installation language of the application.
					   //Q - reserved1 (1 bit): This value is undefined and MUST be ignored.
					   //R - reserved2 (1 bit): This value is undefined and MUST be ignored.
					   //S - fSpare0 (3 bits): This value is undefined and MUST be ignored.
	uint16_t reserved3;//(2 bytes): This value MUST be 0 and MUST be ignored. 
	uint16_t reserved4;//(2 bytes): This value MUST be 0 and MUST be ignored. 
	uint32_t reserved5;//(4 bytes): This value MUST be 0 and MUST be ignored. 
	uint32_t reserved6;//(4 bytes): This value MUST be 0 and MUST be ignored. 
} FibBase;

uint8_t FibBaseA(FibBase *fibBase){
	return fibBase->ABCDEFGHIJKLM & 0x01;
}
uint8_t FibBaseB(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x02) >> 1;
}
uint8_t FibBaseC(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x04) >> 2;
}
uint8_t FibBaseD(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x08) >> 3;
}
uint8_t FibBaseE(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0xF0) >> 4;
}
uint8_t FibBaseF(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x0100) >> 8;
}
uint8_t FibBaseG(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x0200) >> 9;
}
uint8_t FibBaseH(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x0400) >> 10;
}
uint8_t FibBaseI(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x0800) >> 11;
}
uint8_t FibBaseJ(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x1000) >> 12;
}
uint8_t FibBaseK(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x2000) >> 13;
}
uint8_t FibBaseL(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x4000) >> 14;
}
uint8_t FibBaseM(FibBase *fibBase){
	return (fibBase->ABCDEFGHIJKLM & 0x8000) >> 15;
}
uint8_t FibBaseN(FibBase *fibBase){
	return fibBase->NOPQRS & 0x01;
}
uint8_t FibBaseO(FibBase *fibBase){
	return (fibBase->NOPQRS & 0x02) >> 1;
}
uint8_t FibBaseP(FibBase *fibBase){
	return (fibBase->NOPQRS & 0x04) >> 2;
}
uint8_t FibBaseQ(FibBase *fibBase){
	return (fibBase->NOPQRS & 0x08) >> 3;
}
uint8_t FibBaseR(FibBase *fibBase){
	return (fibBase->NOPQRS & 0x10) >> 4;
}
uint8_t FibBaseS(FibBase *fibBase){
	return (fibBase->NOPQRS & 0xE0) >> 5;
}

/*
 * FibRgW97
 * The FibRgW97 structure is a variable-length portion of the Fib.
 */
typedef struct FibRgW97 
{ 
	uint16_t reserved1; //(2 bytes): This value MUST be 0 and MUST be ignored. 
	uint16_t reserved2; //(2 bytes): This value MUST be 0 and MUST be ignored. 
	uint16_t reserved3; //(2 bytes): This value MUST be 0 and MUST be ignored. 
	uint16_t reserved4; //(2 bytes): This value MUST be 0 and MUST be ignored. 
	uint16_t reserved5; //(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
	uint16_t reserved6; //(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
	uint16_t reserved7; //(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
	uint16_t reserved8; //(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
	uint16_t reserved9; //(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
	uint16_t reserved10;//(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
	uint16_t reserved11;//(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
	uint16_t reserved12;//(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
	uint16_t reserved13;//(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
	uint16_t lidFE;     //(2 bytes): A LID whose meaning depends on the nFib value, which is one of the 
						//following:
						//	0x00C1 If FibBase.fFarEast is "true", this is the LID of the stored 
						//		style names. Otherwise it MUST be ignored 
						//	0x00D9, 0x0101, 0x010C, 0x0112 - The LID of the stored style names (STD.
						//		xstzName)

} FibRgW97;

/*
 * FibRgLw97FibRgLw97
 * The FibRgLw97 structure is the third section of the FIB. This contains an array of 4-byte values.
 */
typedef struct FibRgLw97 
{ 
	uint32_t cbMac;     //(4 bytes): Specifies the count of bytes of those written to the WordDocument 
						//stream of the file that have any meaning. All bytes in the WordDocument 
						//stream at offset cbMac and greater MUST be ignored. 
	uint32_t reserved1; //(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t reserved2; //(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t ccpText;   //(4 bytes): A signed integer that specifies the count of CPs in the main 
						//document. This value MUST be zero, 1, or greater. 
	uint32_t ccpFtn;    //(4 bytes): A signed integer that specifies the count of CPs in the footnote 
						//subdocument. This value MUST be zero, 1, or greater 
	uint32_t ccpHdd;    //(4 bytes): A signed integer that specifies the count of CPs in the header 
						//subdocument. This value MUST be zero, 1, or greater 
	uint32_t reserved3; //(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t ccpAtn;    //(4 bytes): A signed integer that specifies the count of CPs in the comment 
						//subdocument. This value MUST be zero, 1, or greater 
	uint32_t ccpEdn;    //(4 bytes): A signed integer that specifies the count of CPs in the endnote 
						//subdocument. This value MUST be zero, 1, or greater 
	uint32_t ccpTxbx;   //(4 bytes): A signed integer that specifies the count of CPs in the textbox 
						//subdocument of the main document. This value MUST be zero, 1, or greater 
	uint32_t ccpHdrTxbx;//(4 bytes): A signed integer that specifies the count of CPs in the textbox 
						//subdocument of the header. This value MUST be zero, 1, or greater 
	uint32_t reserved4; //(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t reserved5; //(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t reserved6; //(4 bytes): This value MUST be equal or less than the number of data elements 
						//in PlcBteChpx, as specified by FibRgFcLcb97.fcPlcfBteChpx and FibRgFcLcb97.
						//lcbPlcfBteChpx. This value MUST be ignored. 
	uint32_t reserved7; //(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t reserved8; //(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t reserved9; //(4 bytes): This value MUST be less than or equal to the number of data 
						//elements in PlcBtePapx, as specified by FibRgFcLcb97.fcPlcfBtePapx and 
						//FibRgFcLcb97.lcbPlcfBtePapx. This value MUST be ignored. 
	uint32_t reserved10;//(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t reserved11;//(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t reserved12;//(4 bytes): This value is undefined and SHOULD be ignored. 
						//Word 97, Word 2000, Word 2002, and Office Word 2003 write a nonzero value 
						//here when saving a document template with changes that require the saving of 
						//an AutoText document.
	uint32_t reserved13;//(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t reserved14;//(4 bytes): This value is undefined and MUST be ignored. 
}FibRgLw97;

/*
 * FibRgFcLcb
 * The FibRgFcLcb structure specifies the file offsets and byte counts for various portions of the data 
 * in the document. The structure of FibRgFcLcb depends on the value of nFib, which is one of the 
 * following.
 */
enum RgFcLcb_t {
	RgFcLcbERROR_t,
	RgFcLcb97_t, 
	RgFcLcb2000_t, 
	RgFcLcb2002_t, 
	RgFcLcb2003_t, 
	RgFcLcb2007_t, 
};

struct nFib2fibRgFcLcb {
	uint16_t nFib;
	enum RgFcLcb_t rgFcLcb;
};

const struct nFib2cbRgFcLcb nFib2fibRgFcLcbTable[] = {
	{0x00C1, RgFcLcb97_t}, 
	{0x00D9, RgFcLcb2000_t}, 
	{0x0101, RgFcLcb2002_t}, 
	{0x010C, RgFcLcb2003_t}, 
	{0x0112, RgFcLcb2007_t}, 
};

static int nFib2fibRgFcLcb_compare(const void *key, const void *value) {
    const struct nFib2fibRgFcLcb *cp1 = key;
    const struct nFib2fibRgFcLcb *cp2 = value;
    return cp1->nFib - cp2->nFib;
}

enum RgFcLcb_t rgFcLcb_get(uint16_t nFib){
    struct nFib2fibRgFcLcb *result = bsearch(&nFib, nFib2fibRgFcLcbTable,
            sizeof(nFib2fibRgFcLcbTable)/sizeof(nFib2fibRgFcLcbTable[0]),
            sizeof(nFib2fibRgFcLcbTable[0]), nFib2fibRgFcLcb_compare);
	if (result)	
		return result->rgFcLcb;
	return RgFcLcbERROR_t;
}

/*
 * FibRgFcLcb97
 * The FibRgFcLcb97 structure is a variable-length portion of the Fib.
 */
typedef struct FibRgFcLcb97 
{ 
	uint32_t fcStshfOrig;  //(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t lcbStshfOrig; //(4 bytes): This value is undefined and MUST be ignored. 
	uint32_t fcStshf;      //(4 bytes): An unsigned integer that specifies an offset in the Table 
	                       //Stream. An STSH that specifies the style sheet for this document 
						   //begins at this offset. 
	uint32_t lcbStshf;     //(4 bytes): An unsigned integer that specifies the size, in bytes, of 
	                       //the STSH that begins at offset fcStshf in the Table Stream. This MUST 
						   //be a nonzero value. 
	uint32_t fcPlcffndRef; //(4 bytes): An unsigned integer that specifies an offset in the Table 
						   //Stream. A PlcffndRef begins at this offset and specifies the locations
						   //of footnote references in the Main Document, and whether those 
						   //references use auto-numbering or custom symbols. If lcbPlcffndRef is 
						   //zero, fcPlcffndRef is undefined and MUST be ignored. 
	uint32_t lcbPlcffndRef;//(4 bytes): An unsigned integer that specifies the size, in bytes, of 
	                       //the PlcffndRef that begins at offset fcPlcffndRef in the Table Stream. 
	uint32_t fcPlcffndTxt; //(4 bytes): An unsigned integer that specifies an offset in the Table 
						   //Stream. A PlcffndTxt begins at this offset and specifies the locations 
						   //of each block of footnote text in the Footnote Document. 
						   //If lcbPlcffndTxt is zero, fcPlcffndTxt is undefined and MUST be 
						   //ignored. 
	uint32_t lcbPlcffndTxt;//(4 bytes):  An unsigned integer that specifies the size, in bytes, 
						   //of the PlcffndTxt that begins at offset fcPlcffndTxt in the Table 
						   //Stream. lcbPlcffndTxt MUST be zero if FibRgLw97.ccpFtn is zero, and 
						   //MUST be nonzero if FibRgLw97.ccpFtn is nonzero.
	uint32_t fcPlcfandRef; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. A PlcfandRef begins at this offset and specifies the dates, 
						   //user initials, and locations of comments in the Main Document. If 
						   //lcbPlcfandRef is zero, fcPlcfandRef is undefined and MUST be ignored.
	uint32_t lcbPlcfandRef;//(4 bytes):  An unsigned integer that specifies the size, in bytes, 
						   //of the PlcfandRef at offset fcPlcfandRef in the Table Stream. 
	uint32_t fcPlcfandTxt; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. A PlcfandTxt begins at this offset and specifies the locations 
						   //of comment text ranges in the Comment Document. If lcbPlcfandTxt is 
						   //zero, fcPlcfandTxt is undefined, and MUST be ignored. 
	uint32_t lcbPlcfandTxt;//(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
						   //the PlcfandTxt at offset fcPlcfandTxt in the Table Stream. 
						   //lcbPlcfandTxt MUST be zero if FibRgLw97.ccpAtn is zero, and MUST be 
						   //nonzero if FibRgLw97.ccpAtn is nonzero.
	uint32_t fcPlcfSed;    //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. A PlcfSed begins at this offset and specifies the locations of
						   //property lists for each section in the Main Document. If lcbPlcfSed is
						   //zero, fcPlcfSed is undefined and MUST be ignored.
	uint32_t lcbPlcfSed;   //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the PlcfSed that begins at offset fcPlcfSed in the Table Stream. 
	uint32_t fcPlcPad;     //(4 bytes):  This value is undefined and MUST be ignored. 
	uint32_t lcbPlcPad;    //(4 bytes):  This value MUST be zero, and MUST be ignored. 
	uint32_t fcPlcfPhe;    //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. A Plc begins at this offset and specifies version-specific 
						   //information about paragraph height. This Plc SHOULD NOT be emitted
						   //and SHOULD be ignored. 
						   //Word 97, Word 2000, and Word 2002 emit this information when 
						   //performing an incremental save.  Office Word 2003, Office Word 2007, 
						   //Word 2010, and Word 2013 do not emit this information
						   //Word 97 reads this information if FibBase.nFib is 193.  Word 2000 
						   //reads this information if FibRgCswNew.nFibNew is 217.  
						   //Word 2002 reads this information if FibRgCswNew.nFibNew is 257.  
						   //Office Word 2003, Office Word 2007, Word 2010, and Word 2013 do not 
						   //read this information
	uint32_t lcbPlcfPhe;   //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
						   //the Plc at offset fcPlcfPhe in the Table Stream. 
	uint32_t fcSttbfGlsy;  //(4 bytes):  An unsigned integer that specifies an offset in the Table
						   //Stream. A SttbfGlsy that contains information about the AutoText items
						   //that are defined in this document begins at this offset. 
	uint32_t lcbSttbfGlsy; //(4 bytes):  An unsigned integer that specifies the size, in bytes, of
						   //the SttbfGlsy at offset fcSttbfGlsy in the Table Stream. If base.fGlsy
						   //of the Fib that contains this FibRgFcLcb97 is zero, this value MUST be
						   //zero.
	uint32_t fcPlcfGlsy;   //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. A PlcfGlsy that contains information about the AutoText items 
						   //that are defined in this document begins at this offset.
	uint32_t lcbPlcfGlsy;  //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
						   //the PlcfGlsy at offset fcPlcfGlsy in the Table Stream. If base.fGlsy 
						   //of the Fib that contains this FibRgFcLcb97 is zero, this value MUST 
						   //be zero. 
	uint32_t fcPlcfHdd;    //(4 bytes):  An unsigned integer that specifies the offset in the Table
						   //Stream where a Plcfhdd begins. The Plcfhdd specifies the locations of 
						   //each block of header/footer text in the WordDocument Stream. If 
						   //lcbPlcfHdd is 0, fcPlcfHdd is undefined and MUST be ignored. 
	uint32_t lcbPlcfHdd;   //(4 bytes):  An unsigned integer that specifies the size, in bytes, of
                           //the Plcfhdd at offset fcPlcfHdd in the Table Stream. If there is no 
						   //Plcfhdd, this value MUST be zero. A Plcfhdd MUST exist if 
						   //FibRgLw97.ccpHdd indicates that there are characters in the Header 
						   //Document (that is, if FibRgLw97.ccpHdd is greater than 0). 
						   //Otherwise, the Plcfhdd MUST NOT exist.
	uint32_t fcPlcfBteChpx;//(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. A PlcBteChpx begins at the offset. fcPlcfBteChpx MUST be 
						   //greater than zero, and MUST be a valid offset in the Table Stream.
	uint32_t lcbPlcfBteChpx;//(4 bytes):  An unsigned integer that specifies the size, in bytes, 
	                       //of the PlcBteChpx at offset fcPlcfBteChpx in the Table Stream. 
						   //lcbPlcfBteChpx MUST be greater than zero. 
	uint32_t fcPlcfBtePapx;//(4 bytes):  An unsigned integer that specifies an offset in the Table
						   //Stream. A PlcBtePapx begins at the offset. fcPlcfBtePapx MUST be 
						   //greater than zero, and MUST be a valid offset in the Table Stream. 
	uint32_t lcbPlcfBtePapx;//(4 bytes):  An unsigned integer that specifies the size, in bytes, of
						   //the PlcBtePapx at offset fcPlcfBtePapx in the Table Stream. 
						   //lcbPlcfBteChpx MUST be greater than zero.
	uint32_t fcPlcfSea;    //(4 bytes):  This value is undefined and MUST be ignored.
	uint32_t lcbPlcfSea;   //(4 bytes):  This value MUST be zero, and MUST be ignored.
	uint32_t fcSttbfFfn;   //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. An SttbfFfn begins at this offset. This table specifies the 
						   //fonts that are used in the document. If lcbSttbfFfn is 0, fcSttbfFfn 
						   //is undefined and MUST be ignored.
	uint32_t lcbSttbfFfn;  //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the SttbfFfn at offset fcSttbfFfn in the Table Stream. 
	uint32_t fcPlcfFldMom; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. A Plcfld begins at this offset and specifies the locations of
						   //field characters in the Main Document. All CPs in this Plcfld MUST be 
						   //greater than or equal to 0 and less than or equal to 
						   //FibRgLw97.ccpText. If lcbPlcfFldMom is zero, fcPlcfFldMom is undefined
						   //and MUST be ignored. 
	uint32_t lcbPlcfFldMom;//(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the Plcfld at offset fcPlcfFldMom in the Table Stream. 
	uint32_t fcPlcfFldHdr; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. A Plcfld begins at this offset and specifies the locations of 
						   //field characters in the Header Document. All CPs in this Plcfld are 
						   //relative to the starting position of the Header Document. All CPs in 
						   //this Plcfld MUST be greater than or equal to zero and less than or 
						   //equal to FibRgLw97.ccpHdd. If lcbPlcfFldHdr is zero, fcPlcfFldHdr is 
						   //undefined and MUST be ignored. 
	uint32_t lcbPlcfFldHdr;//(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
						   //the Plcfld at offset fcPlcfFldHdr in the Table Stream. 
	uint32_t fcPlcfFldFtn; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. A Plcfld begins at this offset and specifies the locations of 
						   //field characters in the Footnote Document. All CPs in this Plcfld are 
						   //relative to the starting position of the Footnote Document. All CPs in 
						   //this Plcfld MUST be greater than or equal to zero and less than or 
						   //equal to FibRgLw97.ccpFtn. If lcbPlcfFldFtn is zero, fcPlcfFldFtn is 
						   //undefined, and MUST be ignored. 
	uint32_t lcbPlcfFldFtn;//(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
						   //the Plcfld at offset fcPlcfFldFtn in the Table Stream 
	uint32_t fcPlcfFldAtn; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. A Plcfld begins at this offset and specifies the locations of 
						   //field characters in the Comment Document. All CPs in this Plcfld are 
						   //relative to the starting position of the Comment Document. All CPs in 
						   //this Plcfld MUST be greater than or equal to zero and less than or 
						   //equal to FibRgLw97.ccpAtn. If lcbPlcfFldAtn is zero, fcPlcfFldAtn is 
						   //undefined and MUST be ignored. 
	uint32_t lcbPlcfFldAtn;//(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the Plcfld at offset fcPlcfFldAtn in the Table Stream 
	uint32_t fcPlcfFldMcr; //(4 bytes):  This value is undefined and MUST be ignored. 
	uint32_t lcbPlcfFldMcr;//(4 bytes):  This value MUST be zero, and MUST be ignored. 
	uint32_t fcSttbfBkmk;  //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. An SttbfBkmk that contains the names of the bookmarks in the 
						   //document begins at this offset. If lcbSttbfBkmk is zero, fcSttbfBkmk 
						   //is undefined and MUST be ignored.  
                           //This SttbfBkmk is parallel to the Plcfbkf at offset fcPlcfBkf in the 
						   //Table Stream. Each string specifies the name of the bookmark that is 
						   //associated with the data element which is located at the same offset 
						   //in that Plcfbkf. For this reason, the SttbfBkmk that begins at offset 
						   //fcSttbfBkmk, and the Plcfbkf that begins at offset fcPlcfBkf, MUST 
						   //contain the same number of elements. 
	uint32_t lcbSttbfBkmk; //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the SttbfBkmk at offset fcSttbfBkmk. 
	uint32_t fcPlcfBkf;    //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. A Plcfbkf that contains information about the standard 
						   //bookmarks in the document begins at this offset. If lcbPlcfBkf is 
						   //zero, fcPlcfBkf is undefined and MUST be ignored. 
                           //Each data element in the Plcfbkf is associated, in a one-to-one 
						   //correlation, with a data element in the Plcfbkl at offset fcPlcfBkl. 
						   //For this reason, the Plcfbkf that begins at offset fcPlcfBkf, and the 
						   //Plcfbkl that begins at offset fcPlcfBkl, MUST contain the same number 
						   //of data elements. This Plcfbkf is parallel to the SttbfBkmk at offset 
						   //fcSttbfBkmk in the Table Stream. Each data element in the Plcfbkf 
						   //specifies information about the bookmark that is associated with the 
						   //element which is located at the same offset in that SttbfBkmk. For 
						   //this reason, the Plcfbkf that begins at offset fcPlcfBkf, and the 
						   //SttbfBkmk that begins at offset fcSttbfBkmk, MUST contain the same 
						   //number of elements. 
						   //The largest value that a CP marking the start or end of a standard 
						   //bookmark is allowed to have is  the CP representing the end of all 
						   //document parts. 
	uint32_t lcbPlcfBkf;   //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the Plcfbkf at offset fcPlcfBkf. 
	uint32_t fcPlcfBkl;    //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. A Plcfbkl that contains information about the standard 
						   //bookmarks in the document begins at this offset. If lcbPlcfBkl is zero,
						   //fcPlcfBkl is undefined and MUST be ignored.  
                           //Each data element in the Plcfbkl is associated, in a one-to-one 
						   //correlation, with a data element in the Plcfbkf at offset fcPlcfBkf. 
						   //For this reason, the Plcfbkl that begins at offset fcPlcfBkl, and the 
						   //Plcfbkf that begins at offset fcPlcfBkf, MUST contain the same number 
						   //of data elements. The largest value that a CP marking the start or end 
						   //of a standard bookmark is allowed to have is the value of the CP 
						   //representing the end of all document parts
	uint32_t lcbPlcfBkl;   //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the Plcfbkl at offset fcPlcfBkl. 
	uint32_t fcCmds;       //(4 bytes):  An unsigned integer that specifies the offset in the Table 
	                       //Stream of a Tcg that specifies command-related customizations. If 
						   //lcbCmds is zero, fcCmds is undefined and MUST be ignored. 
	uint32_t lcbCmds;      //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the Tcg at offset fcCmds. 
	uint32_t fcUnused1;    //(4 bytes):  This value is undefined and MUST be ignored.
	uint32_t lcbUnused1;   //(4 bytes):  This value MUST be zero, and MUST be ignored.
	uint32_t fcSttbfMcr;   //(4 bytes):  This value is undefined and MUST be ignored.
	uint32_t lcbSttbfMcr;  //(4 bytes):  This value MUST be zero, and MUST be ignored.
	uint32_t fcPrDrvr;     //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. The PrDrvr, which contains printer driver information (the 
						   //names of drivers, port, and so on), begins at this offset. If 
						   //lcbPrDrvr is zero, fcPrDrvr is undefined and MUST be ignored.
	uint32_t lcbPrDrvr;    //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the PrDrvr at offset fcPrDrvr. 
	uint32_t fcPrEnvPort;  //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. The PrEnvPort that is the print environment in portrait mode 
						   //begins at this offset. If lcbPrEnvPort is zero, fcPrEnvPort is 
						   //undefined and MUST be ignored. 
	uint32_t lcbPrEnvPort; //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the PrEnvPort at offset fcPrEnvPort 
	uint32_t fcPrEnvLand;  //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. The PrEnvLand that is the print environment in landscape mode 
						   //begins at this offset. If lcbPrEnvLand is zero, fcPrEnvLand is 
						   //undefined and MUST be ignored. 
	uint32_t lcbPrEnvLand; //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the PrEnvLand at offset fcPrEnvLand. 
	uint32_t fcWss;        //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. A Selsf begins at this offset and specifies the last selection 
						   //that was made in the Main Document. If lcbWss is zero, fcWss is 
						   //undefined and MUST be ignored. 
	uint32_t lcbWss;       //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the Selsf at offset fcWss. 
	uint32_t fcDop;        //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. A Dop begins at this offset. 
	uint32_t lcbDop;       //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the Dop at fcDop. This value MUST NOT be zero. 
	uint32_t fcSttbfAssoc; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. An SttbfAssoc that contains strings that are associated with 
						   //the document begins at this offset. 
	uint32_t lcbSttbfAssoc;//(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the SttbfAssoc at offset fcSttbfAssoc. This value MUST NOT be zero 
	uint32_t fcClx;        //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. A Clx begins at this offset. 
	uint32_t lcbClx;       //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the Clx at offset fcClx in the Table Stream. This value MUST be greater 
						   //than zero. 
	uint32_t fcPlcfPgdFtn; //(4 bytes):  This value is undefined and MUST be ignored.
	uint32_t lcbPlcfPgdFtn;//(4 bytes):  This value MUST be zero, and MUST be ignored.
	uint32_t fcAutosaveSource; //(4 bytes):  This value is undefined and MUST be ignored.
	uint32_t lcbAutosaveSource;//(4 bytes):  This value MUST be zero, and MUST be ignored.
	uint32_t fcGrpXstAtnOwners; //(4 bytes):  An unsigned integer that specifies an offset in the 
	                       //Table Stream. An array of XSTs begins at this offset. The value of cch 
						   //for all XSTs in this array MUST be less than 56. The number of entries 
						   //in this array is limited to 0x7FFF. This array contains the names of 
						   //authors of comments in the document. The names in this array MUST be 
						   //unique. If no comments are defined, lcbGrpXstAtnOwners and 
						   //fcGrpXstAtnOwners MUST be zero and MUST be ignored. If any comments 
						   //are in the document, fcGrpXstAtnOwners MUST point to a valid array of 
						   //XSTs. 
	uint32_t lcbGrpXstAtnOwners;//(4 bytes):  An unsigned integer that specifies the size, in bytes, 
	                       //of the XST array at offset fcGrpXstAtnOwners in the Table Stream. 
	uint32_t fcSttbfAtnBkmk; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. An SttbfAtnBkmk that contains information about the annotation 
						   //bookmarks in the document begins at this offset. If lcbSttbfAtnBkmk is 
						   //zero, fcSttbfAtnBkmk is undefined and MUST be ignored.  The 
						   //SttbfAtnBkmk is parallel to the Plcfbkf at offset fcPlcfAtnBkf in the 
						   //Table Stream. Each element in the SttbfAtnBkmk specifies information 
						   //about the bookmark which is associated with the data element that is 
						   //located at the same offset in that Plcfbkf, so the SttbfAtnBkmk 
						   //beginning at offset fcSttbfAtnBkmk and the Plcfbkf beginning at offset 
						   //fcPlcfAtnBkf MUST contain the same number of elements. An additional 
						   //constraint upon the number of elements in the SttbfAtnBkmk is specified 
						   //in the description of fcPlcfAtnBkf. 
	uint32_t lcbSttbfAtnBkmk;//(4 bytes):  An unsigned integer that specifies the size, in bytes, 
	                       //of the SttbfAtnBkmk at offset fcSttbfAtnBkmk. 
	uint32_t fcUnused2;    //(4 bytes):  This value is undefined and MUST be ignored.
	uint32_t lcbUnused2;   //(4 bytes):  This value MUST be zero, and MUST be ignored.
	uint32_t fcUnused3;    //(4 bytes):  This value is undefined and MUST be ignored.
	uint32_t lcbUnused3;   //(4 bytes):  This value MUST be zero, and MUST be ignored.
	uint32_t fcPlcSpaMom;  //(4 bytes):  An unsigned integer that specifies an offset in the Table 
	                       //Stream. A PlcfSpa begins at this offset. The PlcfSpa contains shape 
						   //information for the Main Document. All CPs in this PlcfSpa are relative 
						   //to the starting position of the Main Document and MUST be greater than 
						   //or equal to zero and less than or equal to ccpText in FibRgLw97. 
						   //The final CP is undefined and MUST be ignored, though it MUST be 
						   //greater than the previous entry. If there are no shapes in the Main 
						   //Document, lcbPlcSpaMom and fcPlcSpaMom MUST be zero and MUST be 
						   //ignored. If there are shapes in the Main Document, fcPlcSpaMom MUST 
						   //point to a valid PlcfSpa structure
	uint32_t lcbPlcSpaMom; //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
	                       //the PlcfSpa at offset fcPlcSpaMom. 
	uint32_t fcPlcSpaHdr;  //(4 bytes):  An unsigned integer that specifies an offset in the Table
						   //Stream. A PlcfSpa begins at this offset. The PlcfSpa contains shape 
						   //information for the Header Document. All CPs in this PlcfSpa are 
						   //relative to the starting position of the Header Document and MUST be 
						   //greater than or equal to zero and less than or equal to ccpHdd in 
						   //FibRgLw97. The final CP is undefined and MUST be ignored, though this
						   //value MUST be greater than the previous entry. If there are no shapes
						   //in the Header Document, lcbPlcSpaHdr and fcPlcSpaHdr MUST both be 
						   //zero and MUST be ignored. If there are shapes in the Header Document,
						   //fcPlcSpaHdr MUST point to a valid PlcfSpa structure. 
	uint32_t lcbPlcSpaHdr; //(4 bytes):  An unsigned integer that specifies the size, in bytes, of
						   //the PlcfSpa at the offset fcPlcSpaHdr
	uint32_t fcPlcfAtnBkf; //(4 bytes):  An unsigned integer that specifies an offset in the Table
						   //Stream. A Plcfbkf that contains information about annotation 
						   //bookmarks in the document begins at this offset. If lcbPlcfAtnBkf is 
						   //zero, fcPlcfAtnBkf is undefined and MUST be ignored.
						   //Each data element in the Plcfbkf is associated, in a one-to-one 
						   //correlation, with a data element in the Plcfbkl at offset 
						   //fcPlcfAtnBkl. For this reason, the Plcfbkf that begins at offset 
						   //fcPlcfAtnBkf, and the Plcfbkl that begins at offset fcPlcfAtnBkl, 
						   //MUST contain the same number of data elements. The Plcfbkf is 
						   //parallel to the SttbfAtnBkmk at offset fcSttbfAtnBkmk in the Table 
						   //Stream. Each data element in the Plcfbkf specifies information about 
						   //the bookmark which is associated with the element that is located at 
						   //the same offset in that SttbfAtnBkmk. For this reason, the Plcfbkf 
						   //that begins at offset fcPlcfAtnBkf, and the SttbfAtnBkmk that begins 
						   //at offset fcSttbfAtnBkmk, MUST contain the same number of elements.
						   //The CP range of an annotation bookmark MUST be in the Main Document 
						   //part.
	uint32_t lcbPlcfAtnBkf;//(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
						   //the Plcfbkf at offset fcPlcfAtnBkf.
	uint32_t fcPlcfAtnBkl; //(4 bytes):  An unsigned integer that specifies an offset in the Table
						   //Stream. A Plcfbkl that contains information about annotation 
						   //bookmarks in the document begins at this offset. If lcbPlcfAtnBkl is 
						   //zero, then fcPlcfAtnBkl is undefined and MUST be ignored. 
						   //Each data element in the Plcfbkl is associated, in a one-to-one 
						   //correlation, with a data element in the Plcfbkf at offset 
						   //fcPlcfAtnBkf. For this reason, the Plcfbkl that begins at offset 
						   //fcPlcfAtnBkl, and the Plcfbkf that begins at offset fcPlcfAtnBkf, 
						   //MUST contain the same number of data elements.
						   //The CP range of an annotation bookmark MUST be in the Main Document 
						   //part.
	uint32_t lcbPlcfAtnBkl;//(4 bytes):  An unsigned integer that specifies the size, in bytes, 
						   //of the Plcfbkl at offset fcPlcfAtnBkl.
	uint32_t fcPms;        //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. A Pms, which contains the current state of a print merge 
						   //operation, begins at this offset. If lcbPms is zero, fcPms is 
						   //undefined and MUST be ignored 
	uint32_t lcbPms;       //(4 bytes):  An unsigned integer which specifies the size, in bytes, 
						   //of the Pms at offset fcPms. 
	uint32_t fcFormFldSttbs; //(4 bytes):  This value is undefined and MUST be ignored 
	uint32_t lcbFormFldSttbs;//(4 bytes):  This value MUST be zero, and MUST be ignored 
	uint32_t fcPlcfendRef; //(4 bytes):  An unsigned integer that specifies an offset in the 
						   //Table Stream. A PlcfendRef that begins at this offset specifies the 
						   //locations of endnote references in the Main Document and whether 
						   //those references use auto-numbering or custom symbols. If 
						   //lcbPlcfendRef is zero, fcPlcfendRef is undefined and MUST be ignored. 
	uint32_t lcbPlcfendRef;//(4 bytes):  An unsigned integer that specifies the size, in bytes, 
						   //of the PlcfendRef that begins at offset fcPlcfendRef in the Table 
						   //Stream 
	uint32_t fcPlcfendTxt; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. A PlcfendTxt begins at this offset and specifies the 
						   //locations of each block of endnote text in the Endnote Document. 
						   //If lcbPlcfendTxt is zero, fcPlcfendTxt is undefined and MUST be 
						   //ignored 
	uint32_t lcbPlcfendTxt;//(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
						   //the PlcfendTxt that begins at offset fcPlcfendTxt in the Table Stream.
						   //lcbPlcfendTxt MUST be zero if FibRgLw97.ccpEdn is zero, and MUST be 
						   //nonzero if FibRgLw97.ccpEdn is nonzero. 
	uint32_t fcPlcfFldEdn; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. A Plcfld begins at this offset and specifies the locations of 
						   //field characters in the Endnote Document. All CPs in this Plcfld are 
						   //relative to the starting position of the Endnote Document. All CPs in 
						   //this Plcfld MUST be greater than or equal to zero and less than or 
						   //equal to FibRgLw97.ccpEdn. If lcbPlcfFldEdn is zero, fcPlcfFldEdn is 
						   //undefined and MUST be ignored. 
	uint32_t lcbPlcfFldEdn;//(4 bytes):  An unsigned integer that specifies the size, in bytes, of
						   //the Plcfld at offset fcPlcfFldEdn in the Table Stream 
	uint32_t fcUnused4;    //(4 bytes):  This value is undefined and MUST be ignored 
	uint32_t lcbUnused4;   //(4 bytes):  This value MUST be zero, and MUST be ignored 
	uint32_t fcDggInfo;    //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. An OfficeArtContent that contains information about the 
						   //drawings in the document begins at this offset. 
	uint32_t lcbDggInfo;   //(4 bytes):  An unsigned integer that specifies the size, in bytes, of 
						   //the OfficeArtContent at the offset fcDggInfo. If lcbDggInfo is zero, 
						   //there MUST NOT be any drawings in the document 
	uint32_t fcSttbfRMark; //(4 bytes):  An unsigned integer that specifies an offset in the Table 
						   //Stream. An SttbfRMark that contains the names of authors who have 
						   //added revision marks or comments to the document begins at this 
						   //offset. If lcbSttbfRMark is zero, fcSttbfRMark is undefined and MUST 
						   //be ignored 
	uint32_t lcbSttbfRMark;//(4 bytes):   unsigned integer that specifies the size, in bytes, of 
						   //the SttbfRMark at the offset fcSttbfRMark 
	uint32_t fcSttbfCaption; //(4 bytes):  An unsigned integer that specifies an offset in the 
						   //Table Stream. An SttbfCaption that contains information about the 
						   //captions that are defined in this document begins at this offset. 
						   //If lcbSttbfCaption is zero, fcSttbfCaption is undefined and MUST be 
						   //ignored. If this document is not the Normal template, this value 
						   //MUST be ignored. 
	uint32_t lcbSttbfCaption;//(4 bytes):  An unsigned integer that specifies the size, in bytes, 
						   //of the SttbfCaption at offset fcSttbfCaption in the Table Stream. If 
						   //base.fDot of the Fib that contains this FibRgFcLcb97 is zero, this 
						   //value MUST be zero. 
	uint32_t fcSttbfAutoCaption; //(4 bytes):  An unsigned integer that specifies an offset in the 
						   //Table Stream. A SttbfAutoCaption that contains information about the 
						   //AutoCaption strings defined in this document begins at this offset. 
						   //If lcbSttbfAutoCaption is zero, fcSttbfAutoCaption is undefined and 
						   //MUST be ignored. If this document is not the Normal template, this 
						   //value MUST be ignored
	uint32_t lcbSttbfAutoCaption;//(4 bytes):  An unsigned integer that specifies the size, in 
						   //bytes, of the SttbfAutoCaption at offset fcSttbfAutoCaption in the 
						   //Table Stream. If base.fDot of the Fib that contains this FibRgFcLcb97 
						   //is zero, this MUST be zero 
	uint32_t fcPlcfWkb;    //(4 bytes):   unsigned integer that specifies an offset in the Table 
						   //Stream. A PlcfWKB that contains information about all master 
						   //documents and subdocuments begins at this offset. 
	uint32_t lcbPlcfWkb;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the PlcfWKB at offset fcPlcfWkb in the Table Stream. If lcbPlcfWkb
						   //is zero, fcPlcfWkb is undefined and MUST be ignored 
	uint32_t fcPlcfSpl;    //(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. A Plcfspl, which specifies the state of the spell 
						   //checker for each text range, begins at this offset. If lcbPlcfSpl is 
						   //zero, then fcPlcfSpl is undefined and MUST be ignored. 
	uint32_t lcbPlcfSpl;   //(4 bytes):    unsigned integer that specifies the size, in bytes, of 
						   //the Plcfspl that begins at offset fcPlcfSpl in the Table Stream 
	uint32_t fcPlcftxbxTxt;//(4 bytes):    An unsigned integer that specifies an offset in the 
						   //Table Stream. A PlcftxbxTxt begins at this offset and specifies which 
						   //ranges of text are contained in which textboxes. If lcbPlcftxbxTxt is
						   //zero, fcPlcftxbxTxt is undefined and MUST be ignored. 
	uint32_t lcbPlcftxbxTxt;//(4 bytes):    An unsigned integer that specifies the size, in bytes, 
						   //of the PlcftxbxTxt that begins at offset fcPlcftxbxTxt in the Table 
						   //Stream.
						   //lcbPlcftxbxTxt MUST be zero if FibRgLw97.ccpTxbx is zero, and MUST 
						   //be nonzero if FibRgLw97.ccpTxbx is nonzero. 
	uint32_t fcPlcfFldTxbx;//(4 bytes):    An unsigned integer that specifies an offset in the 
						   //Table Stream. A Plcfld begins at this offset and specifies the 
						   //locations of field characters in the Textbox Document. All CPs in 
						   //this Plcfld are relative to the starting position of the Textbox 
						   //Document. All CPs in this Plcfld MUST be greater than or equal to 
						   //zero and less than or equal to FibRgLw97.ccpTxbx. If lcbPlcfFldTxbx 
						   //is zero, fcPlcfFldTxbx is undefined and MUST be ignored. 
	uint32_t lcbPlcfFldTxbx;//(4 bytes):    An unsigned integer that specifies the size, in bytes, 
						   //of the Plcfld at offset fcPlcfFldTxbx in the Table Stream 
	uint32_t fcPlcfHdrtxbxTxt;//(4 bytes):    An unsigned integer that specifies an offset in the 
						   //Table Stream. A PlcfHdrtxbxTxt begins at this offset and specifies 
						   //which ranges of text are contained in which header textboxes. 
	uint32_t lcbPlcfHdrtxbxTxt;//(4 bytes):    An unsigned integer that specifies the size, in 
						   //bytes, of the PlcfHdrtxbxTxt that begins at offset fcPlcfHdrtxbxTxt 
						   //in the Table Stream.
						   //lcbPlcfHdrtxbxTxt MUST be zero if FibRgLw97.ccpHdrTxbx is zero, and 
						   //MUST be nonzero if FibRgLw97.ccpHdrTxbx is nonzero. 
	uint32_t fcPlcffldHdrTxbx;//(4 bytes):    An unsigned integer that specifies an offset in the 
						   //Table Stream. A Plcfld begins at this offset and specifies the 
						   //locations of field characters in the Header Textbox Document. All CPs
						   //in this Plcfld are relative to the starting position of the Header 
						   //Textbox Document. All CPs in this Plcfld MUST be greater than or 
						   //equal to zero and less than or equal to FibRgLw97.ccpHdrTxbx. If 
						   //lcbPlcffldHdrTxbx is zero, fcPlcffldHdrTxbx is undefined, and MUST be
						   //ignored. 
	uint32_t lcbPlcffldHdrTxbx;//(4 bytes):    An unsigned integer that specifies the size, in 
						   //bytes, of the Plcfld at offset fcPlcffldHdrTxbx in the Table Stream 
	uint32_t fcStwUser;    //(4 bytes):    An unsigned integer that specifies an offset into the 
						   //Table Stream. An StwUser that specifies the user-defined variables 
						   //and VBA digital signature, as specified by [MS-OSHARED] section 
						   //2.3.2, begins at this offset. If lcbStwUser is zero, fcStwUser is 
						   //undefined and MUST be ignored. 
	uint32_t lcbStwUser;   //(4 bytes):    An unsigned integer that specifies the size, in bytes, 
						   //of the StwUser at offset fcStwUser. 
	uint32_t fcSttbTtmbd;  //(4 bytes):    An unsigned integer that specifies an offset into the 
						   //Table Stream. A SttbTtmbd begins at this offset and specifies 
						   //information about the TrueType fonts that are embedded in the 
						   //document. If lcbSttbTtmbd is zero, fcSttbTtmbd is undefined and MUST 
						   //be ignored. 
	uint32_t lcbSttbTtmbd; //(4 bytes):    An unsigned integer that specifies the size, in bytes, 
						   //of the SttbTtmbd at offset fcSttbTtmbd. 
	uint32_t fcCookieData; //(4 bytes):    An unsigned integer that specifies an offset in the 
						   //Table Stream. An RgCdb begins at this offset. If lcbCookieData is 
						   //zero, fcCookieData is undefined and MUST be ignored. Otherwise, 
						   //fcCookieData MAY be ignored. 
						   //Office Word 2007, Word 2010, and Word 2013 ignore this data.
	uint32_t lcbCookieData;//(4 bytes):    An unsigned integer that specifies the size, in bytes, 
						   //of the RgCdb at offset fcCookieData in the Table Stream 
	uint32_t fcPgdMotherOldOld;//(4 bytes):    An unsigned integer that specifies an offset in the
						   //Table Stream. The deprecated document page layout cache begins at 
						   //this offset. Information SHOULD NOT be emitted at this offset and
						   //SHOULD be ignored. If lcbPgdMotherOldOld is zero, fcPgdMotherOldOld 
						   //is undefined and MUST be ignored. 
						   //Word 97 emits information at offset fcPgdMotherOldOld. Neither Word 
						   //2000, Word 2002, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information.
						   //Word 97 reads this information. Word 2000, Word 2002, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information.
	uint32_t lcbPgdMotherOldOld;//(4 bytes):    An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated document page layout cache at offset 
						   //fcPgdMotherOldOld in the Table Stream.
	uint32_t fcBkdMotherOldOld;//(4 bytes):    An unsigned integer that specifies an offset in the
						   //Table Stream. Deprecated document text flow break cache begins at 
						   //this offset. Information SHOULD NOT be emitted at this offset 
						   //and SHOULD be ignored. If lcbBkdMotherOldOld is zero, 
						   //fcBkdMotherOldOld is undefined and MUST be ignored. 
						   //Word 97 emits information at offset fcBkdMotherOldOld. Neither Word 
						   //2000, Word 2002, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information.
						   //Word 97 reads this information. Word 2000, Word 2002, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information.
	uint32_t lcbBkdMotherOldOld;//(4 bytes):    An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated document text flow break cache at offset 
						   //fcBkdMotherOldOld in the Table Stream.
	uint32_t fcPgdFtnOldOld;//(4 bytes):    An unsigned integer that specifies an offset in the 
						   //Table Stream. Deprecated footnote layout cache begins at this 
					       //offset. Information SHOULD NOT be emitted at this offset and 
						   //SHOULD be ignored. If lcbPgdFtnOldOld is zero, fcPgdFtnOldOld is 
						   //undefined and MUST be ignored 
						   //Word 97 emits information at offset fcPgdFtnOldOld. Neither Word 
						   //2000, Word 2002, Office Word 2003, Office Word 2007, Word 2010, 
						   //nor Word 2013 emit this information
						   //Word 97 reads this information. Word 2000, Word 2002, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information.
	uint32_t lcbPgdFtnOldOld;//(4 bytes):    An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated footnote layout cache at offset 
						   //fcPgdFtnOldOld in the Table Stream 
	uint32_t fcBkdFtnOldOld;//(4 bytes):    An unsigned integer that specifies an offset in the 
						   //Table Stream. The deprecated footnote text flow break cache begins 
						   //at this offset. Information SHOULD NOT be emitted at this offset and
						   //SHOULD be ignored. If lcbBkdFtnOldOld is zero, fcBkdFtnOldOld is 
						   //undefined and MUST be ignored. 
						   //Word 97 emits information at offset fcBkdFtnOldOld. Neither Word 
						   //2000, Word 2002, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information.
						   //Word 97 reads this information. Word 2000, Word 2002, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information.
	uint32_t lcbBkdFtnOldOld;//(4 bytes):    An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated footnote text flow break cache at offset 
						   //fcBkdFtnOldOld in the Table Stream. 
	uint32_t fcPgdEdnOldOld;//(4 bytes):    An unsigned integer that specifies an offset in the 
						   //Table Stream. The deprecated endnote layout cache begins at this 
						   //offset. Information SHOULD NOT be emitted at this offset and SHOULD 
						   //be ignored. If lcbPgdEdnOldOld is zero, fcPgdEdnOldOld is undefined 
						   //and MUST be ignored. 
						   //Word 97 emits information at offset fcPgdEdnOldOld. Neither Word 
						   //2000, Word 2002, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information
						   //Word 97 reads this information. Word 2000, Word 2002, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information.
	uint32_t lcbPgdEdnOldOld;//(4 bytes):    An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated endnote layout cache at offset 
						   //fcPgdEdnOldOld in the Table Stream. 
	uint32_t fcBkdEdnOldOld;//(4 bytes):     unsigned integer that specifies an offset in the 
						   //Table Stream. The deprecated endnote text flow break cache begins at 
						   //this offset. Information SHOULD NOT be emitted at this offset and 
						   //SHOULD be ignored. If lcbBkdEdnOldOld is zero, fcBkdEdnOldOld is 
						   //undefined and MUST be ignored. 
						   //Word 97 emits information at offset fcBkdEdnOldOld. Neither Word 
						   //2000, Word 2002, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information.
						   //Only Word 97 reads this information.
	uint32_t lcbBkdEdnOldOld;//(4 bytes):     An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated endnote text flow break cache at offset 
						   //fcBkdEdnOldOld in the Table Stream. 
	uint32_t fcSttbfIntlFld; //(4 bytes):     This value is undefined and MUST be ignored. 
	uint32_t lcbSttbfIntlFld;//(4 bytes):     This value MUST be zero, and MUST be ignored. 
	uint32_t fcRouteSlip;  //(4 bytes):     An unsigned integer that specifies an offset in the 
						   //Table Stream. A RouteSlip that specifies the route slip for this 
						   //document begins at this offset. This value SHOULD be ignored. 
						   //fcRouteSlip is only saved and read by Word 97, Word 2000, Word 2002, 
						   //and Office Word 2003.
	uint32_t lcbRouteSlip; //(4 bytes):     An unsigned integer that specifies the size, in bytes, 
						   //of the RouteSlip at offset fcRouteSlip in the Table Stream 
	uint32_t fcSttbSavedBy;//(4 bytes):     An unsigned integer that specifies an offset in the 
						   //Table Stream. A SttbSavedBy that specifies the save history of this 
						   //document begins at this offset. This value SHOULD be ignored. 
						   //SttbSavedBy is only saved and read by Word 97 and Word 2000.
	uint32_t lcbSttbSavedBy;//(4 bytes):     An unsigned integer that specifies the size, in 
						   //bytes, of the SttbSavedBy at the offset fcSttbSavedBy. This value 
						   //SHOULD be zero 
						   //SttbSavedBy is only saved and read by Word 97 and Word 2000
	uint32_t fcSttbFnm;    //(4 bytes):     An unsigned integer that specifies an offset in the 
						   //Table Stream. An SttbFnm that contains information about the external 
						   //files that are referenced by this document begins at this offset. If 
						   //lcbSttbFnm is zero, fcSttbFnm is undefined and MUST be ignored. 
	uint32_t lcbSttbFnm;   //(4 bytes):     An unsigned integer that specifies the size, in bytes, 
						   //of the SttbFnm at the offset fcSttbFnm. 
	uint32_t fcPlfLst;     //(4 bytes):     An unsigned integer that specifies an offset in the 
						   //Table Stream. A PlfLst that contains list formatting information 
						   //begins at this offset. An array of LVLs is appended to the PlfLst. 
						   //lcbPlfLst does not account for the array of LVLs. The size of the 
						   //array of LVLs is specified by the LSTFs in PlfLst. For each LSTF 
						   //whose fSimpleList is set to 0x1, there is one LVL in the array of 
						   //LVLs that specifies the level formatting of the single level in the 
						   //list which corresponds to the LSTF. And, for each LSTF whose 
						   //fSimpleList is set to 0x0, there are 9 LVLs in the array of LVLs 
						   //that specify the level formatting of the respective levels in the 
						   //list which corresponds to the LSTF. This array of LVLs is in the same
						   //respective order as the LSTFs in PlfLst. If lcbPlfLst is 0, fcPlfLst 
						   //is undefined and MUST be ignored. 
	uint32_t lcbPlfLst;    //(4 bytes):     An unsigned integer that specifies the size, in bytes, 
						   //of the PlfLst at the offset fcPlfLst. This does not include the size 
						   //of the array of LVLs that are appended to the PlfLst 
	uint32_t fcPlfLfo;     //(4 bytes):     An unsigned integer that specifies an offset in the 
						   //Table Stream. A PlfLfo that contains list formatting override 
						   //information begins at this offset. If lcbPlfLfo is zero, fcPlfLfo is 
						   //undefined and MUST be ignored 
	uint32_t lcbPlfLfo;    //(4 bytes):     An unsigned integer that specifies the size, in bytes, 
						   //of the PlfLfo at the offset fcPlfLfo 
	uint32_t fcPlcfTxbxBkd;//(4 bytes):     An unsigned integer that specifies an offset in the 
						   //Table Stream. A PlcfTxbxBkd begins at this offset and specifies which 
						   //ranges of text go inside which textboxes 
	uint32_t lcbPlcfTxbxBkd;//(4 bytes):     An unsigned integer that specifies the size, in 
						   //bytes, of the PlcfTxbxBkd that begins at offset fcPlcfTxbxBkd in the
						   //Table Stream lcbPlcfTxbxBkd MUST be zero if FibRgLw97.ccpTxbx is 
						   //zero, and MUST be nonzero if FibRgLw97.ccpTxbx is nonzero. 
	uint32_t fcPlcfTxbxHdrBkd;//(4 bytes):     An unsigned integer that specifies an offset in the
						   //Table Stream. A PlcfTxbxHdrBkd begins at this offset and specifies 
						   //which ranges of text are contained inside which header textboxes. 
	uint32_t lcbPlcfTxbxHdrBkd;//(4 bytes):     An unsigned integer that specifies the size, in 
						   //bytes, of the PlcfTxbxHdrBkd that begins at offset fcPlcfTxbxHdrBkd 
						   //in the Table Stream.
						   //lcbPlcfTxbxHdrBkd MUST be zero if FibRgLw97.ccpHdrTxbx is zero, and 
						   //MUST be nonzero if FibRgLw97.ccpHdrTxbx is nonzero.
	uint32_t fcDocUndoWord9;//(4 bytes):     An unsigned integer that specifies an offset in the 
						   //WordDocument Stream. Version-specific undo information begins at 
						   //this offset. This information SHOULD NOT be emitted and SHOULD be 
						   //ignored. 
						   //Word 97 and Word 2000 write this information when the user chooses to
						   //save versions in the document. Word 2002, Office Word 2003, Office 
						   //Word 2007, Word 2010, and Word 2013 do not write this information
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 read this 
						   //information. Office Word 2007, Word 2010, and Word 2013 ignore it.
	uint32_t lcbDocUndoWord9;//(4 bytes):     An unsigned integer. If this is nonzero, 
						   //version-specific undo information exists at offset fcDocUndoWord9 in 
						   //the WordDocument Stream 
	uint32_t fcRgbUse;     //(4 bytes):     An unsigned integer that specifies an offset in the 
						   //WordDocument Stream. Version-specific undo information begins at this 
						   //offset. This information SHOULD NOT be emitted and SHOULD be ignored. 
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 write this 
						   //information when the user chooses to save versions in the document. 
						   //Neither Office Word 2007, Word 2010, nor Word 2013 write this 
						   //information.
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 read this 
						   //information. Office Word 2007, Word 2010, and Word 2013 ignore it.
	uint32_t lcbRgbUse;    //(4 bytes):      unsigned integer that specifies the size, in bytes, 
						   //of the version-specific undo information at offset fcRgbUse in the 
						   //WordDocument Stream 
	uint32_t fcUsp;        //(4 bytes):      An unsigned integer that specifies an offset in the 
						   //WordDocument Stream. Version- specific undo information begins at 
						   //this offset. This information SHOULD NOT be emitted and SHOULD be
						   //ignored. 
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 write this 
						   //information when the user chooses to save versions in the document. 
						   //Neither Office Word 2007, Word 2010, nor Word 2013 write this 
						   //information.
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 read this 
						   //information. Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbUsp;       //(4 bytes):      An unsigned integer that specifies the size, in 
						   //bytes, of the version-specific undo information at offset fcUsp in 
						   //the WordDocument Stream. 
	uint32_t fcUskf;       //(4 bytes):      An unsigned integer that specifies an offset in the 
						   //Table Stream. Version-specific undo information begins at this 
						   //offset. This information SHOULD NOT be emitted and SHOULD be ignored. 
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 write this 
						   //information when the user chooses to save versions in the document. 
						   //Neither Office Word 2007, Word 2010, nor Word 2013 write this 
						   //information.
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 read this 
						   //information. Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbUskf;      //(4 bytes):       unsigned integer that specifies the size, in bytes, 
						   //of the version-specific undo information at offset fcUskf in the 
						   //Table Stream 
	uint32_t fcPlcupcRgbUse;//(4 bytes):       An unsigned integer that specifies an offset in the
						   //Table Stream. A Plc begins at this offset and contains 
						   //version-specific undo information. This information SHOULD NOT be 
						   //emitted and SHOULD be ignored 
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 write this 
						   //information when the user chooses to save versions in the document. 
						   //Neither Office Word 2007, Word 2010, nor Word 2013 write this 
						   //information.
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 read this 
						   //information. Office Word 2007, Word 2010, and wd15 ignore it
	uint32_t lcbPlcupcRgbUse;//(4 bytes):       An unsigned integer that specifies the size, in 
						   //bytes, of the Plc at offset fcPlcupcRgbUse in the Table Stream
	uint32_t fcPlcupcUsp;  //(4 bytes):       An unsigned integer that specifies an offset in the 
						   //Table Stream. A Plc begins at this offset and contains 
						   //version-specific undo information. This information SHOULD NOT be 
						   //emitted and SHOULD be ignored. 
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 write this 
						   //information when the user chooses to save versions in the document. 
						   //Neither Office Word 2007, Word 2010, nor Word 2013 write this 
						   //information
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 read this 
						   //information. Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbPlcupcUsp; //(4 bytes):       An unsigned integer that specifies the size, in 
						   //bytes, of the Plc at offset fcPlcupcUsp in the Table Stream 
	uint32_t fcSttbGlsyStyle;//(4 bytes):       An unsigned integer that specifies an offset in 
						   //the Table Stream. A SttbGlsyStyle, which contains information about 
						   //the styles that are used by the AutoText items which are defined in 
						   //this document, begins at this offset. 
	uint32_t lcbSttbGlsyStyle;//(4 bytes):       An unsigned integer that specifies the size, in 
						   //bytes, of the SttbGlsyStyle at offset fcSttbGlsyStyle in the Table 
						   //Stream. If base.fGlsy of the Fib that contains this FibRgFcLcb97 is 
						   //zero, this value MUST be zero. 
	uint32_t fcPlgosl;     //(4 bytes):       An unsigned integer that specifies an offset in the 
						   //Table Stream. A PlfGosl begins at the offset. If lcbPlgosl is zero, 
						   //fcPlgosl is undefined and MUST be ignored 
	uint32_t lcbPlgosl;    //(4 bytes):       An unsigned integer that specifies the size, in 
						   //bytes, of the PlfGosl at offset fcPlgosl in the Table Stream 
	uint32_t fcPlcocx;     //(4 bytes):       An unsigned integer that specifies an offset in the 
						   //Table Stream. A RgxOcxInfo that specifies information about the OLE 
						   //controls in the document begins at this offset. When there are no OLE 
						   //controls in the document, fcPlcocx and lcbPlcocx MUST be zero and 
						   //MUST be ignored. If there are any OLE controls in the document, 
						   //fcPlcocx MUST point to a valid RgxOcxInfo 
	uint32_t lcbPlcocx;    //(4 bytes):       An unsigned integer that specifies the size, in 
						   //bytes, of the RgxOcxInfo at the offset fcPlcocx 
	uint32_t fcPlcfBteLvc; //(4 bytes):       An unsigned integer that specifies an offset in the 
						   //Table Stream. A deprecated numbering field cache begins at this 
						   //offset. This information SHOULD NOT be emitted and SHOULD ignored. If 
						   //lcbPlcfBteLvc is zero, fcPlcfBteLvc is undefined and MUST be ignored 
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 write this 
						   //information. Neither Office Word 2007, Word 2010, nor Word 2013 write
						   //this information
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 read this 
						   //information. Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbPlcfBteLvc;//(4 bytes):       An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated numbering field cache at offset fcPlcfBteLvc
						   //in the Table Stream. This value SHOULD<58> be zero 
	uint32_t dwLowDateTime;//(4 bytes):       The low-order part of a FILETIME structure, as 
						   //specified by [MS- DTYP], that specifies when the document was last 
						   //saved 
	uint32_t dwHighDateTime;//(4 bytes):       The high-order part of a FILETIME structure, as 
						   //specified by [MS- DTYP], that specifies when the document was last 
						   //saved 
	uint32_t fcPlcfLvcPre10;//(4 bytes):       An unsigned integer that specifies an offset in the
						   //Table Stream. The deprecated list level cache begins at this offset.
						   //Information SHOULD NOT be emitted at this offset and SHOULD be 
						   //ignored. If lcbPlcfLvcPre10 is zero, fcPlcfLvcPre10 is undefined and 
						   //MUST be ignored. 
						   //Word 97 emits information at offset fcPlcfLvcPre10 when performing an
						   //incremental save. Word 2000 emits information at offset 
						   //fcPlcfLvcPre10 on every save. Neither Word 2002, Office Word 2003, 
						   //Office Word 2007, Word 2010, nor Word 2013 emit information at offset
						   //fcPlcfLvcPre10 and the value of fcPlcfLvcPre10 is undefined
						   //Word 97 and Word 2000 read this information. Word 2002, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbPlcfLvcPre10;//(4 bytes):       An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated list level cache at offset fcPlcfLvcPre10 in
						   //the Table Stream. This value SHOULD be zero
						   //Word 97 and Word 2000 write lcbPlcfLvcPre10 with the size, in bytes, 
						   //of the information emitted at offset fcPlcfLvcPre10. Word 2002, 
						   //Office Word 2003, Office Word 2007, Word 2010, and Word 2013 write 0 
						   //to lcbPlcfLvcPre10
	uint32_t fcPlcfAsumy;  //(4 bytes):       An unsigned integer that specifies an offset in the 
						   //Table Stream. A PlcfAsumy begins at the offset. If lcbPlcfAsumy is 
						   //zero, fcPlcfAsumy is undefined and MUST be ignored 
	uint32_t lcbPlcfAsumy; //(4 bytes):        unsigned integer that specifies the size, in bytes,
						   //of the PlcfAsumy at offset fcPlcfAsumy in the Table Stream 
	uint32_t fcPlcfGram;   //(4 bytes):       An unsigned integer that specifies an offset in the
						   //Table Stream. A Plcfgram, which specifies the state of the grammar 
						   //checker for each text range, begins at this offset. If lcbPlcfGram is
						   //zero, then fcPlcfGram is undefined and MUST be ignored
	uint32_t lcbPlcfGram;  //(4 bytes):       An unsigned integer that specifies the size, in 
						   //bytes, of the Plcfgram that begins at offset fcPlcfGram in the Table 
						   //Stream
	uint32_t fcSttbListNames;//(4 bytes):     An unsigned integer that specifies an offset in the 
						   //Table Stream. A SttbListNames, which specifies the LISTNUM field 
						   //names of the lists in the document, begins at this offset. If 
						   //lcbSttbListNames is zero, fcSttbListNames is undefined and MUST be 
						   //ignored 
	uint32_t lcbSttbListNames;//(4 bytes):     An unsigned integer that specifies the size, in 
						   //bytes, of the SttbListNames at the offset fcSttbListNames 
	uint32_t fcSttbfUssr;  //(4 bytes):     An unsigned integer that specifies an offset in the 
						   //Table Stream. The deprecated, version-specific undo information 
						   //begins at this offset. This information SHOULD NOT be emitted and 
						   //SHOULD be ignored 
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 write this 
						   //information when the user chooses to save versions in the document. 
						   //Neither Office Word 2007, Word 2010, nor Word 2013 write this 
						   //information
						   //Word 97, Word 2000, Word 2002, and Office Word 2003 read this 
						   //information. Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbSttbfUssr; //(4 bytes):     An unsigned integer that specifies the size, in bytes,
						   //of the deprecated, version-specific undo information at offset 
						   //fcSttbfUssr in the Table Stream 
} FibRgFcLcb97;

/*
 * FibRgFcLcb2000
 * The FibRgFcLcb2000 structure is a variable-sized portion of the Fib. It extends the 
 * FibRgFcLcb97.
 */
typedef struct FibRgFcLcb2000 
{
	FibRgFcLcb97 rgFcLcb97;//(744 bytes): The contained FibRgFcLcb97
	uint32_t fcPlcfTch;    //(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. A PlcfTch begins at this offset and specifies a cache 
						   //of table characters. Information at this offset SHOULD be ignored. If 
						   //lcbPlcfTch is zero, fcPlcfTch is undefined and MUST be ignored 
	uint32_t lcbPlcfTch;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the PlcfTch at offset fcPlcfTch 
	uint32_t fcRmdThreading;//(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. An RmdThreading that specifies the data concerning the 
						   //e-mail messages and their authors in this document begins at this 
						   //offset 
	uint32_t lcbRmdThreading;//(4 bytes):   An unsigned integer that specifies the size, in bytes,
						   //of the RmdThreading at the offset fcRmdThreading. This value MUST NOT
						   //be zero. 
	uint32_t fcMid;        //(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. A double-byte character Unicode string that specifies 
						   //the message identifier of the document begins at this offset. This 
						   //value MUST be ignored 
	uint32_t lcbMid;       //(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the double-byte character Unicode string at offset fcMid. This 
						   //value MUST be ignored 
	uint32_t fcSttbRgtplc; //(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. A SttbRgtplc that specifies the styles of lists in the 
						   //document begins at this offset. If lcbSttbRgtplc is zero, 
						   //fcSttbRgtplc is undefined and MUST be ignored 
	uint32_t lcbSttbRgtplc;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the SttbRgtplc at the offset fcSttbRgtplc 
	uint32_t fcMsoEnvelope;//(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. An MsoEnvelopeCLSID, which specifies the envelope data
						   //as specified by [MS-OSHARED] section 2.3.8.1, begins at this offset. 
						   //If lcbMsoEnvelope is zero, fcMsoEnvelope is undefined and MUST be 
						   //ignored 
	uint32_t lcbMsoEnvelope;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the MsoEnvelopeCLSID at the offset fcMsoEnvelope 
	uint32_t fcPlcfLad;    //(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. A Plcflad begins at this offset and specifies the 
						   //language auto-detect state of each text range. If lcbPlcfLad is zero, 
						   //fcPlcfLad is undefined and MUST be ignored 
	uint32_t lcbPlcfLad;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the Plcflad that begins at offset fcPlcfLad in the Table Stream 
	uint32_t fcRgDofr;     //(4 bytes):   An unsigned integer that specifies an offset in the Table
						   //Stream. A variable-length array with elements of type Dofrh begins at 
						   //that offset. The elements of this array are records that support the 
						   //frame set and list style features. If lcbRgDofr is zero, fcRgDofr is 
						   //undefined and MUST be ignored. 
	uint32_t lcbRgDofr;    //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the array that begins at offset fcRgDofr in the Table Stream.
	uint32_t fcPlcosl;     //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A PlfCosl begins at the offset. If lcbPlcosl is zero, fcPlcosl
						   //is undefined and MUST be ignored.
	uint32_t lcbPlcosl;    //(4 bytes):   An unsigned integer that specifies the size, in bytes, of 
	                       //the PlfCosl at offset fcPlcosl in the Table Stream.
	uint32_t fcPlcfCookieOld;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. A PlcfcookieOld begins at this offset. If 
						   //lcbPlcfcookieOld is zero, fcPlcfcookieOld is undefined and MUST be 
						   //ignored. fcPlcfcookieOld MAY be ignored. 
						   //Word 2002, Office Word 2003, Office Word 2007, Word 2010, and Word 
						   //2013 ignore this value.
	uint32_t lcbPlcfCookieOld;//(4 bytes):   An unsigned integer that specifies the size, in bytes,
                           //of the PlcfcookieOld at offset fcPlcfcookieOld in the Table Stream. 
	uint32_t fcPgdMotherOld;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. The deprecated document page layout cache begins at this
						   //offset. Information SHOULD NOT be emitted at this offset and SHOULD be
						   //ignored. If lcbPgdMotherOld is zero, fcPgdMotherOld is undefined and 
						   //MUST be ignored.
						   //Word 2000 and Word 2002 emit information at offset fcPgdMotherOld. 
						   //Neither Word 97, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information
						   //Word 2000 and Word 2002 read this information. Word 97, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information
	uint32_t lcbPgdMotherOld;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the deprecated document page layout cache at offset fcPgdMotherOld 
						   //in the Table Stream 
	uint32_t fcBkdMotherOld;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. The deprecated document text flow break cache begins at 
						   //this offset. Information SHOULD NOT be emitted at this offset and 
						   //SHOULD be ignored. If lcbBkdMotherOld is zero, fcBkdMotherOld is 
						   //undefined and MUST be ignored 
						   //Word 2000 and Word 2002 emit information at offset fcBkdMotherOld. 
						   //Neither Word 97, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information
						   //Word 2000 and Word 2002 read this information. Word 97, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information
	uint32_t lcbBkdMotherOld;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the deprecated document text flow break cache at offset 
						   //fcBkdMotherOld in the Table Stream 
	uint32_t fcPgdFtnOld;  //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. The deprecated footnote layout cache begins at this offset. 
						   //Information SHOULD NOT be emitted at this offset and SHOULD be 
						   //ignored. If lcbPgdFtnOld is zero, fcPgdFtnOld is undefined and MUST be
						   //ignored 
						   //Word 2000 and Word 2002 emit information at offset fcPgdFtnOld. 
						   //Neither Word 97, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information
						   //Word 2000 and Word 2002 read this information. Word 97, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information
	uint32_t lcbPgdFtnOld; //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the deprecated footnote layout cache at offset fcPgdFtnOld in the 
						   //Table Stream
	uint32_t fcBkdFtnOld;  //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. The deprecated footnote text flow break cache begins at this 
						   //offset. Information SHOULD NOT be emitted at this offset and SHOULD be
						   //ignored. If lcbBkdFtnOld is zero, fcBkdFtnOld is undefined and MUST be
						   //ignored
						   //Word 2000 and Word 2002 emit information at offset fcBkdFtnOld. 
						   //Neither Word 97, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information
						   //Word 2000 and Word 2002 read this information. Word 97, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information
	uint32_t lcbBkdFtnOld; //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the deprecated footnote text flow break cache at offset fcBkdFtnOld 
						   //in the Table Stream
	uint32_t fcPgdEdnOld;  //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. The deprecated endnote layout cache begins at this offset. 
						   //Information SHOULD NOT be emitted at this offset and SHOULD be 
						   //ignored. If lcbPgdEdnOld is zero, fcPgdEdnOld is undefined and MUST be
						   //ignored
						   //Word 2000 and Word 2002 emit information at offset fcPgdEdnOld. 
						   //Neither Word 97, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information
						   //Word 2000 and Word 2002 read this information. Word 97, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information
	uint32_t lcbPgdEdnOld; //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the deprecated endnote layout cache at offset fcPgdEdnOld in the Table
						   //Stream
	uint32_t fcBkdEdnOld;  //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. The deprecated endnote text flow break cache begins at this 
						   //offset. Information SHOULD NOT be emitted at this offset and SHOULD be
						   //ignored. If lcbBkdEdnOld is zero, fcBkdEdnOld is undefined and MUST be
						   //ignored
						   //Word 2000 and Word 2002 emit information at offset fcBkdEdnOld. 
						   //Neither Word 97, Office Word 2003, Office Word 2007, Word 2010, nor 
						   //Word 2013 emit this information
						   //Word 2000 and Word 2002 read this information. Word 97, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information
	uint32_t lcbBkdEdnOld; //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the deprecated endnote text flow break cache at offset fcBkdEdnOld in 
						   //the Table Stream
} FibRgFcLcb2000;

/*
 * FibRgFcLcb2002
 * The FibRgFcLcb2002 structure is a variable-sized portion of the Fib. It extends the 
 * FibRgFcLcb2000.
 */
typedef struct FibRgFcLcb2002 
{
	FibRgFcLcb2000 rgFcLcb2000;//(864 bytes):  The contained FibRgFcLcb2000.
	uint32_t fcUnused1;    //(4 bytes):   This value is undefined and MUST be ignored.
	uint32_t lcbUnused1;   //(4 bytes):   This value MUST be zero, and MUST be ignored.
	uint32_t fcPlcfPgp;    //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A PGPArray begins at this offset. If lcbPlcfPgp is 0, 
						   //fcPlcfPgp is undefined and MUST be ignored
	uint32_t lcbPlcfPgp;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the PGPArray that is stored at offset fcPlcfPgp
	uint32_t fcPlcfuim;    //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A Plcfuim begins at this offset. If lcbPlcfuim is zero, 
						   //fcPlcfuim is undefined and MUST be ignored
	uint32_t lcbPlcfuim;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the Plcfuim at offset fcPlcfuim
	uint32_t fcPlfguidUim; //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A PlfguidUim begins at this offset. If lcbPlfguidUim is zero, 
						   //fcPlfguidUim is undefined and MUST be ignored
	uint32_t lcbPlfguidUim;//(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the PlfguidUim at offset fcPlfguidUim
	uint32_t fcAtrdExtra;  //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. An AtrdExtra begins at this offset. If lcbAtrdExtra is zero, 
						   //fcAtrdExtra is undefined and MUST be ignored
	uint32_t lcbAtrdExtra; //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the AtrdExtra at offset fcAtrdExtra in the Table Stream
	uint32_t fcPlrsid;     //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A PLRSID begins at this offset. If lcbPlrsid is zero, fcPlrsid
						   //is undefined and MUST be ignored
	uint32_t lcbPlrsid;    //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the PLRSID at offset fcPlrsid in the Table Stream
	uint32_t fcSttbfBkmkFactoid;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. An SttbfBkmkFactoid containing information about smart 
						   //tag bookmarks in the document begins at this offset. If 
						   //lcbSttbfBkmkFactoid is zero, fcSttbfBkmkFactoid is undefined and MUST 
						   //be ignored.  
						   //The SttbfBkmkFactoid is parallel to the Plcfbkfd at offset 
						   //fcPlcfBkfFactoid in the Table Stream. Each element in the 
						   //SttbfBkmkFactoid specifies information about the bookmark that is 
						   //associated with the data element which is located at the same offset 
						   //in that Plcfbkfd. For this reason, the SttbfBkmkFactoid that begins at
						   //offset fcSttbfBkmkFactoid, and the Plcfbkfd that begins at offset 
						   //fcPlcfBkfFactoid, MUST contain the same number of elements
	uint32_t lcbSttbfBkmkFactoid;//(4 bytes):   An unsigned integer that specifies the size, in 
	                       //bytes, of the SttbfBkmkFactoid at offset fcSttbfBkmkFactoid 
	uint32_t fcPlcfBkfFactoid;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. A Plcfbkfd that contains information about the smart tag
						   //bookmarks in the document begins at this offset. If lcbPlcfBkfFactoid 
						   //is zero, fcPlcfBkfFactoid is undefined and MUST be ignored.  
						   //Each data element in the Plcfbkfd is associated, in a one-to-one 
						   //correlation, with a data element in the Plcfbkld at offset 
						   //fcPlcfBklFactoid. For this reason, the Plcfbkfd that begins at offset 
						   //fcPlcfBkfFactoid, and the Plcfbkld that begins at offset 
						   //fcPlcfBklFactoid, MUST contain the same number of data elements. The 
						   //Plcfbkfd is parallel to the SttbfBkmkFactoid at offset 
						   //fcSttbfBkmkFactoid in the Table Stream. Each data element in the 
						   //Plcfbkfd specifies information about the bookmark that is associated 
						   //with the element which is located at the same offset in that 
						   //SttbfBkmkFactoid. For this reason, the Plcfbkfd that begins at offset 
						   //fcPlcfBkfFactoid, and the SttbfBkmkFactoid that begins at offset 
						   //fcSttbfBkmkFactoid, MUST contain the same number of elements
	uint32_t lcbPlcfBkfFactoid;//(4 bytes):   An unsigned integer that specifies the size, in 
	                       //bytes, of the Plcfbkfd at offset fcPlcfBkfFactoid 
	uint32_t fcPlcfcookie; //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A Plcfcookie begins at this offset. If lcbPlcfcookie is zero, 
						   //fcPlcfcookie is undefined and MUST be ignored. fcPlcfcookie MAY be 
						   //ignored 
						   //Office Word 2007, Word 2010, and Word 2013 ignore this value.
	uint32_t lcbPlcfcookie;//(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the Plcfcookie at offset fcPlcfcookie in the Table Stream
	uint32_t fcPlcfBklFactoid;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. A Plcfbkld that contains information about the smart tag
						   //bookmarks in the document begins at this offset. If lcbPlcfBklFactoid 
						   //is zero, fcPlcfBklFactoid is undefined and MUST be ignored.  
						   //Each data element in the Plcfbkld is associated, in a one-to-one 
						   //correlation, with a data element in the Plcfbkfd at offset 
						   //fcPlcfBkfFactoid. For this reason, the Plcfbkld that begins at offset 
						   //fcPlcfBklFactoid, and the Plcfbkfd that begins at offset 
						   //fcPlcfBkfFactoid, MUST contain the same number of data elements
	uint32_t lcbPlcfBklFactoid;//(4 bytes):   An unsigned integer that specifies the size, in 
	                       //bytes, of the Plcfbkld at offset fcPlcfBklFactoid 
	uint32_t fcFactoidData;//(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A SmartTagData begins at this offset and specifies information
						   //about the smart tag recognizers that are used in this document. 
						   //If lcbFactoidData is zero, fcFactoidData is undefined and MUST be 
						   //ignored 
	uint32_t lcbFactoidData;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the SmartTagData at offset fcFactoidData in the Table Stream
	uint32_t fcDocUndo;    //(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //WordDocument Stream. Version-specific undo information begins at this 
						   //offset. This information SHOULD NOT be emitted and SHOULD be ignored 
						   //Word 2002 and Office Word 2003 write this information when the user 
						   //chooses to save versions in the document.  Neither Word 97, Word 2000,
						   //Office Word 2007, Word 2010, nor Word 2013 write this information
						   //Word 2002 and Office Word 2003 read this information.  Word 97, Word 
						   //2000, Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbDocUndo;   //(4 bytes):   An unsigned integer. If this value is nonzero, 
	                       //version-specific undo information exists at offset fcDocUndo in the 
						   //WordDocument Stream 
	uint32_t fcSttbfBkmkFcc;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. An SttbfBkmkFcc that contains information about the 
						   //format consistency-checker bookmarks in the document begins at this 
						   //offset. If lcbSttbfBkmkFcc is zero, fcSttbfBkmkFcc is undefined and 
						   //MUST be ignored.  
						   //The SttbfBkmkFcc is parallel to the Plcfbkfd at offset fcPlcfBkfFcc in
						   //the Table Stream. Each element in the SttbfBkmkFcc specifies 
						   //information about the bookmark that is associated with the data 
						   //element which is located at the same offset in that Plcfbkfd. For this
						   //reason, the SttbfBkmkFcc that begins at offset fcSttbfBkmkFcc, and the
						   //Plcfbkfd that begins at offset fcPlcfBkfFcc, MUST contain the same 
						   //number of elements 
	uint32_t lcbSttbfBkmkFcc;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the SttbfBkmkFcc at offset fcSttbfBkmkFcc 
	uint32_t fcPlcfBkfFcc; //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A Plcfbkfd that contains information about format 
						   //consistency-checker bookmarks in the document begins at this offset. 
						   //If lcbPlcfBkfFcc is zero, fcPlcfBkfFcc is undefined and MUST be 
						   //ignored 
						   //Each data element in the Plcfbkfd is associated, in a one-to-one 
						   //correlation, with a data element in the Plcfbkld at offset 
						   //fcPlcfBklFcc. For this reason, the Plcfbkfd that begins at offset 
						   //fcPlcfBkfFcc and the Plcfbkld that begins at offset fcPlcfBklFcc MUST 
						   //contain the same number of data elements. The Plcfbkfd is parallel to 
						   //the SttbfBkmkFcc at offset fcSttbfBkmkFcc in the Table Stream. Each 
						   //data element in the Plcfbkfd specifies information about the bookmark 
						   //that is associated with the element which is located at the same 
						   //offset in that SttbfBkmkFcc. For this reason, the Plcfbkfd that begins
						   //at offset fcPlcfBkfFcc and the SttbfBkmkFcc that begins at offset 
						   //fcSttbfBkmkFcc MUST contain the same number of elements
	uint32_t lcbPlcfBkfFcc;//(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the Plcfbkfd at offset fcPlcfBkfFcc
	uint32_t fcPlcfBklFcc; //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A Plcfbkld that contains information about the format 
						   //consistency-checker bookmarks in the document begins at this offset. 
						   //If lcbPlcfBklFcc is zero, fcPlcfBklFcc is undefined and MUST be 
						   //ignored.  
						   //Each data element in the Plcfbkld is associated, in a one-to-one 
						   //correlation, with a data element in the Plcfbkfd at offset 
						   //fcPlcfBkfFcc. For this reason, the Plcfbkld that begins at offset 
						   //fcPlcfBklFcc, and the Plcfbkfd that begins at offset fcPlcfBkfFcc, 
						   //MUST contain the same number of data elements
	uint32_t lcbPlcfBklFcc;//(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the Plcfbkld at offset fcPlcfBklFcc
	uint32_t fcSttbfbkmkBPRepairs;//(4 bytes):   An unsigned integer that specifies an offset in 
	                       //the Table Stream. An SttbfBkmkBPRepairs that contains information 
						   //about the repair bookmarks in the document begins at this offset. 
						   //If lcbSttbfBkmkBPRepairs is zero, fcSttbfBkmkBPRepairs is undefined 
						   //and MUST be ignored. 
						   //The SttbfBkmkBPRepairs is parallel to the Plcfbkf at offset 
						   //fcPlcfBkfBPRepairs in the Table Stream. Each element in the 
						   //SttbfBkmkBPRepairs specifies information about the bookmark that is 
						   //associated with the data element which is located at the same offset 
						   //in that Plcfbkf. For this reason, the SttbfBkmkBPRepairs that begins 
						   //at offset fcSttbfBkmkBPRepairs, and the Plcfbkf that begins at offset 
						   //fcPlcfBkfBPRepairs, MUST contain the same number of elements
	uint32_t lcbSttbfbkmkBPRepairs;//(4 bytes):   An unsigned integer that specifies the size, in 
	                       //bytes, of the SttbfBkmkBPRepairs at offset fcSttbfBkmkBPRepairs 
	uint32_t fcPlcfbkfBPRepairs;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. A Plcfbkf that contains information about the repair 
						   //bookmarks in the document begins at this offset. If 
						   //lcbPlcfBkfBPRepairs is zero, fcPlcfBkfBPRepairs is undefined and MUST 
						   //be ignored.  
						   //Each data element in the Plcfbkf is associated, in a one-to-one 
						   //correlation, with a data element in the Plcfbkl at offset 
						   //fcPlcfBklBPRepairs. For this reason, the Plcfbkf that begins at offset 
						   //fcPlcfBkfBPRepairs, and the Plcfbkl that begins at offset 
						   //fcPlcfBklBPRepairs, MUST contain the same number of data elements. 
						   //The Plcfbkf is parallel to the SttbfBkmkBPRepairs at offset 
						   //fcSttbfBkmkBPRepairs in the Table Stream. Each data element in the 
						   //Plcfbkf specifies information about the bookmark that is associated 
						   //with the element which is located at the same offset in that 
						   //SttbfBkmkBPRepairs. For this reason, the Plcfbkf that begins at offset 
						   //fcPlcfbkfBPRepairs, and the SttbfBkmkBPRepairs that begins at offset 
						   //fcSttbfBkmkBPRepairs, MUST contain the same number of elements. The 
						   //CPs in this Plcfbkf MUST NOT exceed the CP that represents the end of 
						   //the Main Document part 
	uint32_t lcbPlcfbkfBPRepairs;//(4 bytes):   An unsigned integer that specifies the size, in 
	                       //bytes, of the Plcfbkf at offset fcPlcfbkfBPRepairs 
	uint32_t fcPlcfbklBPRepairs;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. A Plcfbkl that contains information about the repair 
						   //bookmarks in the document begins at this offset. If 
						   //lcbPlcfBklBPRepairs is zero, fcPlcfBklBPRepairs is undefined and MUST 
						   //be ignored. 
						   //Each data element in the Plcfbkl is associated, in a one-to-one 
						   //correlation, with a data element in the Plcfbkf at offset 
						   //fcPlcfBkfBPRepairs. For this reason, the Plcfbkl that begins at offset 
						   //fcPlcfBklBPRepairs, and the Plcfbkf that begins at offset 
						   //fcPlcfBkfBPRepairs, MUST contain the same number of data elements.  
						   //The CPs that are contained in this Plcfbkl MUST NOT exceed the CP that
						   //represents the end of the Main Document part 
	uint32_t lcbPlcfbklBPRepairs;//(4 bytes):   An unsigned integer that specifies the size, in 
	                       //bytes, of the Plcfbkl at offset fcPlcfBklBPRepairs 
	uint32_t fcPmsNew;     //(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. A new Pms, which contains the current state of a print 
						   //merge operation, begins at this offset. If lcbPmsNew is zero, 
						   //fcPmsNew is undefined and MUST be ignored 
	uint32_t lcbPmsNew;    //(4 bytes):   An unsigned integer which specifies the size, in bytes, 
	                       //of the Pms at offset fcPmsNew. 
	uint32_t fcODSO;       //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. Office Data Source Object (ODSO) data that is used to perform 
						   //mail merge begins at this offset. The data is stored in an array of 
						   //ODSOPropertyBase items. The ODSOPropertyBase items are of variable 
						   //size and are stored contiguously. The complete set of properties that
						   //are contained in the array is determined by reading each 
						   //ODSOPropertyBase, until a total of lcbODSO bytes of data are read. If 
						   //lcbODSO is zero, fcODSO is undefined and MUST be ignored 
	uint32_t lcbODSO;      //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the Office Data Source Object data at offset fcODSO in the Table 
						   //Stream
	uint32_t fcPlcfpmiOldXP;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. The deprecated paragraph mark information cache begins 
						   //at this offset. Information SHOULD NOT be emitted at this offset and 
						   //SHOULD be ignored. If lcbPlcfpmiOldXP is zero, fcPlcfpmiOldXP is 
						   //undefined and MUST be ignored
						   //Word 2002 emits information at offset fcPlcfpmiOldXP. Neither Word 97,
						   //Word 2000, Office Word 2003, Office Word 2007, Word 2010, nor Word 
						   //2013 emit information at this offset and the value of fcPlcfpmiOldXP 
						   //is undefined
						   //Word 2002 reads this information. Word 97, Word 2000, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbPlcfpmiOldXP;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the deprecated paragraph mark information cache at offset 
						   //fcPlcfpmiOldXP in the Table Stream. This value SHOULD be zero. 
						   //Word 2002 writes lcbPlcfpmiOldXP with the size, in bytes, of the 
						   //information emitted at offset fcPlcfpmiOldXP. Office Word 2003, Office
						   //Word 2007, Word 2010, and Word 2013 write 0 to lcbPlcfpmiOldXP. 
						   //Neither Word 97 nor Word 2000 write a FibRgFcLcb2002.
	uint32_t fcPlcfpmiNewXP;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. The deprecated paragraph mark information cache begins 
						   //at this offset. Information SHOULD NOT be emitted at this offset and 
						   //SHOULD be ignored. If lcbPlcfpmiNewXP is zero, fcPlcfpmiNewXP is 
						   //undefined and MUST be ignored 
						   //Word 2002 emits information at offset fcPlcfpmiNewXP. Neither Word 97, 
						   //Word 2000, Office Word 2003, Office Word 2007, Word 2010, nor Word 
						   //2013 emit information at this offset and the value of fcPlcfpmiNewXP 
						   //is undefined
						   //Word 2002 reads this information. Word 97, Word 2000, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbPlcfpmiNewXP;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the deprecated paragraph mark information cache at offset 
						   //fcPlcfpmiNewXP in the Table Stream. This value SHOULD be zero
						   //Word 2002 writes lcbPlcfpmiNewXP with the size, in bytes, of the 
						   //information emitted at offset fcPlcfpmiNewXP. Office Word 2003, Office
						   //Word 2007, Word 2010, and Word 2013 write 0 to lcbPlcfpmiNewXP. 
						   //Neither Word 97 nor Word 2000 write a FibRgFcLcb2002.
	uint32_t fcPlcfpmiMixedXP;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. The deprecated paragraph mark information cache begins 
						   //at this offset. Information SHOULD NOT be emitted at this offset and 
						   //SHOULD be ignored. If lcbPlcfpmiMixedXP is zero, fcPlcfpmiMixedXP is 
						   //undefined and MUST be ignored 
						   //Word 2002 emits information at offset fcPlcfpmiMixedXP. Neither Word 
						   //97, Word 2000, Office Word 2003, Office Word 2007, Word 2010, nor Word
						   //2013 emit information at this offset and the value of fcPlcfpmiMixedXP
						   //is undefined
						   //Word 2002 reads this information. Word 97, Word 2000, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbPlcfpmiMixedXP;//(4 bytes):   An unsigned integer that specifies the size, in 
	                       //bytes, of the deprecated paragraph mark information cache at offset 
						   //fcPlcfpmiMixedXP in the Table Stream. This value SHOULD be zero 
						   //Word 2002 writes lcbPlcfpmiMixedXP with the size, in bytes, of the 
						   //information emitted at offset fcPlcfpmiMixedXP. Office Word 2003, 
						   //Office Word 2007, Word 2010, and Word 2013 write 0 to 
						   //lcbPlcfpmiMixedXP. Neither Word 97 nor Word 2000 write a 
						   //FibRgFcLcb2002.
	uint32_t fcUnused2;    //(4 bytes):   This value is undefined and MUST be ignored. 
	uint32_t lcbUnused2;   //(4 bytes):   This value MUST be zero, and MUST be ignored. 
	uint32_t fcPlcffactoid;//(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A Plcffactoid, which specifies the smart tag recognizer state 
						   //of each text range, begins at this offset. If lcbPlcffactoid is zero, 
						   //fcPlcffactoid is undefined and MUST be ignored. 
	uint32_t lcbPlcffactoid;//(4 bytes):   An unsigned integer that specifies the size, in bytes of
						   //the Plcffactoid that begins at offset fcPlcffactoid in the Table 
						   //Stream.
	uint32_t fcPlcflvcOldXP;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. The deprecated listnum field cache begins at this 
						   //offset. Information SHOULD NOT be emitted at this offset and SHOULD be
						   //ignored. If lcbPlcflvcOldXP is zero, fcPlcflvcOldXP is undefined and 
						   //MUST be ignored
						   //Word 2002 emits information at offset fcPlcflvcOldXP. Neither Word 97,
						   //Word 2000, Office Word 2003, Office Word 2007, Word 2010, nor Word 
						   //2013 emit information at this offset and the value of fcPlcflvcOldXP 
						   //is undefined.
						   //Word 2002 reads this information. Word 97, Word 2000, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbPlcflvcOldXP;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the deprecated listnum field cache at offset fcPlcflvcOldXP in the 
						   //Table Stream. This value SHOULD be zero 
						   //Word 2002 writes lcbPlcflvcOldXP with the size, in bytes, of the 
						   //information emitted at offset fcPlcflvcOldXP. Office Word 2003, Office
						   //Word 2007, Word 2010, and Word 2013 write 0 to lcbPlcflvcOldXP. 
						   //Neither Word 97 nor Word 2000 write a FibRgFcLcb2002.
	uint32_t fcPlcflvcNewXP;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. The deprecated listnum field cache begins at this 
						   //offset. Information SHOULD NOT be emitted at this offset and SHOULD be
						   //ignored. If lcbPlcflvcNewXP is zero, fcPlcflvcNewXP is undefined and 
						   //MUST be ignored. 
						   //Word 2002 emits information at offset fcPlcflvcNewXP. Neither Word 97,
						   //Word 2000, Office Word 2003, Office Word 2007, Word 2010, nor Word 
						   //2013 emit information at this offset and the value of fcPlcflvcNewXP 
						   //is undefined
						   //Word 2002 reads this information. Word 97, Word 2000, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbPlcflvcNewXP;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the deprecated listnum field cache at offset fcPlcflvcNewXP in the 
						   //Table Stream. This value SHOULD be zero. 
						   //Word 2002 writes lcbPlcflvcNewXP with the size, in bytes, of the 
						   //information emitted at offset fcPlcflvcNewXP. Office Word 2003, Office
						   //Word 2007, Word 2010, and Word 2013 write 0 to lcbPlcflvcNewXP. 
						   //Neither Word 97 nor Word 2000 write a FibRgFcLcb2002.
	uint32_t fcPlcflvcMixedXP;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. The deprecated listnum field cache begins at this 
						   //offset. Information SHOULD NOT be emitted at this offset and SHOULD be
						   //ignored. If lcbPlcflvcMixedXP is zero, fcPlcflvcMixedXP is undefined 
						   //and MUST be ignored 
						   //Word 2002 emits information at offset fcPlcflvcMixedXP. Neither Word 
						   //97, Word 2000, Office Word 2003, Office Word 2007, Word 2010, nor Word
						   //2013 emit information at this offset and the value of fcPlcflvcMixedXP
						   //is undefined
						   //Word 2002 reads this information. Word 97, Word 2000, Office Word 
						   //2003, Office Word 2007, Word 2010, and Word 2013 ignore it
	uint32_t lcbPlcflvcMixedXP;//(4 bytes):   An unsigned integer that specifies the size, in 
	                       //bytes, of the deprecated listnum field cache at offset 
						   //fcPlcflvcMixedXP in the Table Stream. This value SHOULD be zero 
						   //Word 2002 writes lcbPlcflvcMixedXP with the size, in bytes, of the 
						   //information emitted at offset fcPlcflvcMixedXP. Office Word 2003, 
						   //Office Word 2007, Word 2010, and Word 2013 write 0 to 
						   //lcbPlcflvcMixedXP. Neither Word 97 nor Word 2000 write a 
						   //FibRgFcLcb2002.
} FibRgFcLcb2002;

/*
 * FibRgFcLcb2003
 * The FibRgFcLcb2003 structure is a variable-sized portion of the Fib.  It extends the 
 * FibRgFcLcb2002.
 */
typedef struct FibRgFcLcb2003 
{
	FibRgFcLcb2002 rgFcLcb2002;//(1088 bytes): The contained FibRgFcLcb2002.
	uint32_t fcHplxsdr;    //(4 bytes):   An unsigned integer that specifies an offset in the Table 
	                       //Stream. An Hplxsdr structure begins at this offset. This structure 
						   //specifies information about XML schema definition (XSD) references 
	uint32_t lcbHplxsdr;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, of 
	                       //the Hplxsdr structure at the offset fcHplxsdr in the Table Stream. If 
						   //lcbHplxsdr is zero, then fcHplxsdr is undefined and MUST be ignored 
	uint32_t fcSttbfBkmkSdt;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. An SttbfBkmkSdt that contains information about the 
						   //structured document tag bookmarks in the document begins at this 
						   //offset. If lcbSttbfBkmkSdt is zero, then fcSttbfBkmkSdt is undefined 
						   //and MUST be ignored.  The SttbfBkmkSdt is parallel to the Plcbkfd at 
						   //offset fcPlcfBkfSdt in the Table Stream. Each element in the 
						   //SttbfBkmkSdt specifies information about the bookmark that is 
						   //associated with the data element which is located at the same offset in
						   //that Plcbkfd. For this reason, the SttbfBkmkSdt that begins at offset 
						   //fcSttbfBkmkSdt, and the Plcbkfd that begins at offset fcPlcfBkfSdt, 
						   //MUST contain the same number of elements 
	uint32_t lcbSttbfBkmkSdt;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the SttbfBkmkSdt at offset fcSttbfBkmkSdt 
	uint32_t fcPlcfBkfSdt; //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A Plcbkfd that contains information about the structured 
						   //document tag bookmarks in the document begins at this offset. If 
						   //lcbPlcfBkfSdt is zero, fcPlcfBkfSdt is undefined and MUST be ignored.  
						   //Each data element in the Plcbkfd is associated, in a one-to-one 
						   //correlation, with a data element in the Plcbkld at offset 
						   //fcPlcfBklSdt. For this reason, the Plcbkfd that begins at offset 
						   //fcPlcfBkfSdt, and the Plcbkld that begins at offset fcPlcfBklSdt, MUST
						   //contain the same number of data elements. The Plcbkfd is parallel to 
						   //the SttbfBkmkSdt at offset fcSttbfBkmkSdt in the Table Stream. Each 
						   //data element in the Plcbkfd specifies information about the bookmark 
						   //that is associated with the element which is located at the same 
						   //offset in that SttbfBkmkSdt. For this reason, the Plcbkfd that begins
						   //at offset fcPlcfBkfSdt, and the SttbfBkmkSdt that begins at offset 
						   //fcSttbfBkmkSdt, MUST contain the same number of elements 
	uint32_t lcbPlcfBkfSdt;//(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the Plcbkfd at offset fcPlcfBkfSdt.
	uint32_t fcPlcfBklSdt; //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A Plcbkld that contains information about the structured 
						   //document tag bookmarks in the document begins at this offset. If 
						   //lcbPlcfBklSdt is zero, fcPlcfBklSdt is undefined and MUST be ignored.  
						   //Each data element in the Plcbkld is associated, in a one-to-one 
						   //correlation, with a data element in the Plcbkfd at offset 
						   //fcPlcfBkfSdt. For this reason, the Plcbkld that begins at offset 
						   //fcPlcfBklSdt, and the Plcbkfd that begins at offset fcPlcfBkfSdt MUST 
						   //contain the same number of data elements.
	uint32_t lcbPlcfBklSdt;//(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the Plcbkld at offset fcPlcfBklSdt.
	uint32_t fcCustomXForm;//(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. An array of 16-bit Unicode characters, which specifies the 
						   //full path and file name of the XML Stylesheet to apply when saving 
						   //this document in XML format, begins at this offset. If lcbCustomXForm 
						   //is zero, fcCustomXForm is undefined and MUST be ignored.
	uint32_t lcbCustomXForm;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the array at offset fcCustomXForm in the Table Stream. This value 
						   //MUST be less than or equal to 4168 and MUST be evenly divisible by 
						   //two.
	uint32_t fcSttbfBkmkProt;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. An SttbfBkmkProt that contains information about 
						   //range-level protection bookmarks in the document begins at this 
						   //offset. If lcbSttbfBkmkProt is zero, fcSttbfBkmkProt is undefined and 
						   //MUST be ignored.  
						   //The SttbfBkmkProt is parallel to the Plcbkf at offset fcPlcfBkfProt 
						   //in the Table Stream. Each element in the SttbfBkmkProt specifies 
						   //information about the bookmark that is associated with the data 
						   //element which is located at the same offset in that Plcbkf. For this 
						   //reason, the SttbfBkmkProt that begins at offset fcSttbfBkmkProt, and 
						   //the Plcbkf that begins at offset fcPlcfBkfProt, MUST contain the same 
						   //number of elements. 
	uint32_t lcbSttbfBkmkProt;//(4 bytes):   An unsigned integer that specifies the size, in bytes,
                           //of the SttbfBkmkProt at offset fcSttbfBkmkProt. 
	uint32_t fcPlcfBkfProt;//(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. A Plcbkf that contains information about range-level 
						   //protection bookmarks in the document begins at this offset. If 
						   //lcbPlcfBkfProt is zero, then fcPlcfBkfProt is undefined and MUST be 
						   //ignored. 
						   //Each data element in the Plcbkf is associated, in a one-to-one 
						   //correlation, with a data element in the Plcbkl at offset 
						   //fcPlcfBklProt. For this reason, the Plcbkf that begins at offset 
						   //fcPlcfBkfProt, and the Plcbkl that begins at offset fcPlcfBklProt, 
						   //MUST contain the same number of data elements. The Plcbkf is parallel 
						   //to the SttbfBkmkProt at offset fcSttbfBkmkProt in the Table Stream. 
						   //Each data element in the Plcbkf specifies information about the 
						   //bookmark that is associated with the element which is located at the 
						   //same offset in that SttbfBkmkProt. For this reason, the Plcbkf that 
						   //begins at offset fcPlcfBkfProt, and the SttbfBkmkProt that begins at 
						   //offset fcSttbfBkmkProt, MUST contain the same number of elements.
	uint32_t lcbPlcfBkfProt;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the Plcbkf at offset fcPlcfBkfProt.
	uint32_t fcPlcfBklProt;//(4 bytes):   An unsigned integer that specifies an offset in the Table
						   //Stream. A Plcbkl containing information about range-level protection 
						   //bookmarks in the document begins at this offset. If lcbPlcfBklProt is 
						   //zero, then fcPlcfBklProt is undefined and MUST be ignored.  
						   //Each data element in the Plcbkl is associated in a one-to-one 
						   //correlation with a data element in the Plcbkf at offset fcPlcfBkfProt,
						   //so the Plcbkl beginning at offset fcPlcfBklProt and the Plcbkf 
						   //beginning at offset fcPlcfBkfProt MUST contain the same number of data
						   //elements.
	uint32_t lcbPlcfBklProt;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the Plcbkl at offset fcPlcfBklProt.
	uint32_t fcSttbProtUser;//(4 bytes):   An unsigned integer that specifies an offset in the 
	                       //Table Stream. A SttbProtUser that specifies the usernames that are 
						   //used for range-level protection begins at this offset. 
	uint32_t lcbSttbProtUser;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
	                       //of the SttbProtUser at the offset fcSttbProtUser. 
	uint32_t fcUnused;     //(4 bytes):   This value is undefined and MUST be ignored. 
	uint32_t lcbUnused;    //(4 bytes):   This value MUST be zero, and MUST be ignored. 
	uint32_t fcPlcfpmiOld; //(4 bytes):   An unsigned integer that specifies an offset in the Table
                           //Stream. Deprecated paragraph mark information cache begins at this 
						   //offset. Information SHOULD NOT be emitted at this offset and SHOULD be
						   //ignored. If lcbPlcfpmiOld is zero, then fcPlcfpmiOld is undefined and 
						   //MUST be ignored. 
						   //Only Office Word 2003 emits information at offset fcPlcfpmiOld; 
						   //Neither Office Word 2007, Word 2010, nor Word 2013 emit information 
						   //at this offset and the value of fcPlcfpmiOld is undefined
						   //Only Office Word 2003 reads this information.
	uint32_t lcbPlcfpmiOld;//(4 bytes):   An unsigned integer that specifies the size, in bytes, of
                           //the deprecated paragraph mark information cache at offset fcPlcfpmiOld
						   //in the Table Stream. SHOULD be zero.
						   //Office Word 2003 writes lcbPlcfpmiOld with the size, in bytes, of the 
						   //information emitted at offset fcPlcfpmiOld; Office Word 2007, Word 
						   //2010, and Word 2013 write 0 to lcbPlcfpmiOld.
	uint32_t fcPlcfpmiOldInline;//(4 bytes):   An unsigned integer that specifies an offset in the
						   //Table Stream. Deprecated paragraph mark information cache begins at 
						   //this offset. Information SHOULD NOT be emitted at this offset and 
						   //SHOULD be ignored. If lcbPlcfpmiOldInline is zero, then 
						   //fcPlcfpmiOldInline is undefined and MUST be ignored.
						   //Only Office Word 2003 emits information at offset fcPlcfpmiOldInline;
						   //Neither Office Word 2007, Word 2010, nor Word 2013 emit information 
						   //at this offset and the value of fcPlcfpmiOldInline is undefined
						   //Only Office Word 2003 reads this information
	uint32_t lcbPlcfpmiOldInline;//(4 bytes):   An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated paragraph mark information cache at offset 
						   //fcPlcfpmiOldInline in the Table Stream. SHOULD be zero
						   //Office Word 2003 writes lcbPlcfpmiOldInline with the size, in bytes, 
						   //of the information emitted at offset fcPlcfpmiOldInline; Office Word 
						   //2007, Word 2010, and Word 2013 write 0 to lcbPlcfpmiOldInline
	uint32_t fcPlcfpmiNew; //(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. Deprecated paragraph mark information cache begins at 
						   //this offset. Information SHOULD NOT be emitted at this offset and 
						   //SHOULD be ignored. If lcbPlcfpmiNew is zero, then fcPlcfpmiNew is 
						   //undefined and MUST be ignored. 
						   //Only Office Word 2003 emits information at offset fcPlcfpmiNew; 
						   //Neither Office Word 2007, Word 2010, nor Word 2013 emit information 
						   //at this offset and the value of fcPlcfpmiNew is undefined
						   //Only Office Word 2003 reads this information
	uint32_t lcbPlcfpmiNew;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the deprecated paragraph mark information cache at offset 
						   //fcPlcfpmiNew in the Table Stream. SHOULD be zero. 
						   //Office Word 2003 writes lcbPlcfpmiNew with the size, in bytes, of the
						   //information emitted at offset fcPlcfpmiNew; Office Word 2007, Word 
						   //2010, and Word 2013 write 0 to lcbPlcfpmiNew
	uint32_t fcPlcfpmiNewInline;//(4 bytes):   An unsigned integer that specifies an offset in the
						   //Table Stream. Deprecated paragraph mark information cache begins at 
						   //this offset. Information SHOULD NOT be emitted at this offset and
						   //SHOULD be ignored. If lcbPlcfpmiNewInline is zero, then 
						   //fcPlcfpmiNewInline is undefined and MUST be ignored 
						   //Only Office Word 2003 emits information at offset fcPlcfpmiNewInline;
						   //Neither Office Word 2007, Word 2010, nor Word 2013 emit information
						   //at this offset and the value of fcPlcfpmiNewInline is undefined
						   //Only Office Word 2003 reads this information
	uint32_t lcbPlcfpmiNewInline;//(4 bytes):   An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated paragraph mark information cache at offset 
						   //fcPlcfpmiNewInline in the Table Stream. SHOULD be zero.
						   //Office Word 2003 writes lcbPlcfpmiNewInline with the size, in bytes, 
						   //of the information emitted at offset fcPlcfpmiNewInline; Office Word 
						   //2007, Word 2010, and Word 2013 write 0 to lcbPlcfpmiNewInline
	uint32_t fcPlcflvcOld; //(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. Deprecated listnum field cache begins at this offset. 
						   //Information SHOULD NOT be emitted at this offset and SHOULD be 
						   //ignored. If lcbPlcflvcOld is zero, then fcPlcflvcOld is undefined 
						   //and MUST be ignored
						   //Only Office Word 2003 emits information at offset fcPlcflvcOld; 
						   //Neither Office Word 2007, Word 2010, nor Word 2013 emit information 
						   //at this offset and the value of fcPlcflvcOld is undefined.
						   //Only Office Word 2003 reads this information
	uint32_t lcbPlcflvcOld;//(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the deprecated listnum field cache at offset fcPlcflvcOld in the 
						   //Table Stream. SHOULD be zero 
						   //Office Word 2003 writes lcbPlcflvcOld with the size, in bytes, of the
						   //information emitted at offset fcPlcflvcOld; Office Word 2007, Word 
						   //2010, and Word 2013 write 0 to lcbPlcflvcOld
	uint32_t fcPlcflvcOldInline;//(4 bytes):   An unsigned integer that specifies an offset in the
						   //Table Stream. Deprecated listnum field cache begins at this offset. 
						   //Information SHOULD NOT be emitted at this offset and SHOULD be 
						   //ignored. If lcbPlcflvcOldInline is zero, fcPlcflvcOldInline is 
						   //undefined and MUST be ignored 
						   //Only Office Word 2003 emits information at offset fcPlcflvcOldInline;
						   //Neither Office Word 2007, Word 2010, nor Word 2013 emit information 
						   //at this offset and the value of fcPlcflvcOldInline is undefined
						   //Only Office Word 2003 reads this information
	uint32_t lcbPlcflvcOldInline;//(4 bytes):   An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated listnum field cache at offset 
						   //fcPlcflvcOldInline in the Table Stream. SHOULD be zero
						   //Office Word 2003 writes lcbPlcflvcOldInline with the size, in bytes,
						   //of the information emitted at offset fcPlcflvcOldInline; Office Word
						   //2007, Word 2010, and Word 2013 write 0 to lcbPlcflvcOldInline
	uint32_t fcPlcflvcNew; //(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. Deprecated listnum field cache begins at this offset.
						   //Information SHOULD NOT be emitted at this offset and SHOULD be 
						   //ignored. If lcbPlcflvcNew is zero, fcPlcflvcNew is undefined and 
						   //MUST be ignored 
						   //Only Office Word 2003 emits information at offset fcPlcflvcNew; 
						   //Neither Office Word 2007, Word 2010, nor Word 2013 emit information 
						   //at this offset and the value of fcPlcflvcNew is undefined
						   //Only Office Word 2003 reads this information
	uint32_t lcbPlcflvcNew; //(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the deprecated listnum field cache at offset fcPlcflvcNew in the 
						   //Table Stream. SHOULD be zero 
						   //Office Word 2003 writes lcbPlcflvcNew with the size, in bytes, of the
						   //information emitted at offset fcPlcflvcNew; Office Word 2007, Word 
						   //2010, and Word 2013 write 0 to lcbPlcflvcNew
	uint32_t fcPlcflvcNewInline; //(4 bytes):   An unsigned integer that specifies an offset in 
						   //the Table Stream. Deprecated listnum field cache begins at this 
						   //offset. Information SHOULD NOT be emitted at this offset and SHOULD 
						   //be ignored. If lcbPlcflvcNewInline is zero, fcPlcflvcNewInline is 
						   //undefined and MUST be ignored 
						   //Only Office Word 2003 emits information at offset fcPlcflvcNewInline;
						   //Neither Office Word 2007, Word 2010, nor Word 2013 emit information
						   //at this offset and the value of fcPlcflvcNewInline is undefined
						   //Only Office Word 2003 reads this information
	uint32_t lcbPlcflvcNewInline; //(4 bytes):   An unsigned integer that specifies the size, in 
						   //bytes, of the deprecated listnum field cache at offset 
						   //fcPlcflvcNewInline in the Table Stream. SHOULD be zero 
						   //Office Word 2003 writes lcbPlcflvcNewInline with the size, in bytes, 
						   //of the information emitted at offset fcPlcflvcNewInline; Office Word 
						   //2007, Word 2010, and Word 2013 write 0 to lcbPlcflvcNewInline
	uint32_t fcPgdMother;  //(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. Deprecated document page layout cache begins at this 
						   //offset. Information SHOULD NOT be emitted at this offset and 
						   //SHOULD be ignored. If lcbPgdMother is zero, fcPgdMother is 
						   //undefined and MUST be ignored 
						   //Office Word 2003 emits information at offset fcPgdMother. Neither 
						   //Word 97, Word 2000, Office Word 2003, Office Word 2007, Word 2010, 
						   //nor Word 2013 emit this information
						   //Office Word 2003 reads this information. Word 97, Word 2000, Word 
						   //2002, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information
	uint32_t lcbPgdMother; //(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the deprecated document page layout cache at offset fcPgdMother 
						   //in the Table Stream 
	uint32_t fcBkdMother;  //(4 bytes):   An unsigned integer that specifies an offset in the 
						   //Table Stream. Deprecated document text flow break cache begins at 
						   //this offset. Information SHOULD NOT be emitted at this offset and 
						   //SHOULD be ignored. If lcbBkdMother is zero, then fcBkdMother is 
						   //undefined and MUST be ignored 
						   //Office Word 2003 emits information at offset fcBkdMother. Neither 
						   //Word 97, Word 2000, Word 2002, Office Word 2007, Word 2010, nor Word 
						   //2013 emit this information
						   //Office Word 2003 reads this information. Word 97, Word 2000, Word 
						   //2002, Office Word 2007, Word 2010, and Word 2013 ignore this 
						   //information
	uint32_t lcbBkdMother; //(4 bytes):   An unsigned integer that specifies the size, in bytes, 
						   //of the deprecated document text flow break cache at offset 
						   //fcBkdMother in the Table Stream 
	uint32_t fcAfdMother; //(4 bytes):   An unsigned integer that specifies an offset in the Table
						  //Stream. Deprecated document author filter cache begins at this offset.
						  //Information SHOULD NOT be emitted at this offset and SHOULD be 
						  //ignored. If lcbAfdMother is zero, then fcAfdMother is undefined and 
						  //MUST be ignored 
						  //Office Word 2003 emits information at offset fcAfdMother. Neither Word
						  //97, Word 2000, Word 2002, Office Word 2007, Word 2010, nor Word 2013 
						  //emit this information
						  //Office Word 2003 reads this information. Word 97, Word 2000, Word 
						  //2002, Office Word 2007, Word 2010, and Word 2013 ignore this 
						  //information
	uint32_t lcbAfdMother;//(4 bytes):   An unsigned integer that specifies the size, in bytes, of
						  //the deprecated document author filter cache at offset fcAfdMother in 
						  //the Table Stream
	uint32_t fcPgdFtn;    //(4 bytes):   An unsigned integer that specifies an offset in the Table
						  //Stream. Deprecated footnote layout cache begins at this offset. 
						  //Information SHOULD NOT be emitted at this offset and SHOULD be 
						  //ignored. If lcbPgdFtn is zero, then fcPgdFtn is undefined and MUST be 
						  //ignored
						  //Office Word 2003 emits information at offset fcPgdFtn. Neither Word 
						  //97, Word 2000, Word 2002, Office Word 2007, Word 2010, nor Word 2013 
						  //emit this information
						  //Office Word 2003 reads this information. Word 97, Word 2000, Word 
						  //2002, Office Word 2007, Word 2010, and Word 2013 ignore this 
						  //information
	uint32_t lcbPgdFtn;   //(4 bytes):   unsigned integer that specifies the size, in bytes, of 
						  //the deprecated footnote layout cache at offset fcPgdFtn in the Table 
						  //Stream
	uint32_t fcBkdFtn;    //(4 bytes):   An unsigned integer that specifies an offset in the Table
						  //Stream. The deprecated footnote text flow break cache begins at this
						  //offset. Information SHOULD NOT be emitted at this offset and SHOULD be
						  //ignored. If lcbBkdFtn is zero, fcBkdFtn is undefined and MUST be 
						  //ignored 
						  //Office Word 2003 emits information at offset fcBkdFtn. Neither Word 
						  //97, Word 2000, Word 2002, Office Word 2007, Word 2010, nor Word 2013 
						  //emit this information
						  //Office Word 2003 reads this information. Word 97, Word 2000, Word 
						  //2002, Office Word 2007, Word 2010, and Word 2013 ignore this 
						  //information
	uint32_t lcbBkdFtn;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
						  //the deprecated footnote text flow break cache at offset fcBkdFtn in 
						  //the Table Stream
	uint32_t fcAfdFtn;    //(4 bytes):   An unsigned integer that specifies an offset in the Table
						  //Stream. The deprecated footnote author filter cache begins at this 
						  //offset. Information SHOULD NOT be emitted at this offset and SHOULD be
						  //ignored. If lcbAfdFtn is zero, fcAfdFtn is undefined and MUST be 
						  //ignored
						  //Office Word 2003 emits information at offset fcAfdFtn. Neither Word 
						  //97, Word 2000, Word 2002, Office Word 2007, Word 2010, nor Word 2013 
						  //emit this information
						  //Office Word 2003 reads this information. Word 97, Word 2000, Word 
						  //2002, Office Word 2007, Word 2010, and Word 2013 ignore this 
						  //information
	uint32_t lcbAfdFtn;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
						  //the deprecated footnote author filter cache at offset fcAfdFtn in the
						  //Table Stream
	uint32_t fcPgdEdn;    //(4 bytes):   An unsigned integer that specifies an offset in the Table
						  //Stream. The deprecated endnote layout cache begins at this offset. 
						  //Information SHOULD NOT be emitted at this offset and SHOULD be 
						  //ignored. If lcbPgdEdn is zero, then fcPgdEdn is undefined and MUST be 
						  //ignored
						  //Office Word 2003 emits information at offset fcPgdEdn. Neither Word 
						  //97, Word 2000, Word 2002, Office Word 2007, Word 2010, nor Word 2013 
						  //emit this information
						  //Office Word 2003 reads this information. Word 97, Word 2000, Word 
						  //2002, Office Word 2007, Word 2010, and Word 2013 ignore this 
						  //information
	uint32_t lcbPgdEdn;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
						  //the deprecated endnote layout cache at offset fcPgdEdn in the Table 
						  //Stream
	uint32_t fcBkdEdn;    //(4 bytes):   An unsigned integer that specifies an offset in the Table
						  //Stream. The deprecated endnote text flow break cache begins at this 
						  //offset. Information SHOULD NOT be emitted at this offset and SHOULD be
						  //ignored. If lcbBkdEdn is zero, fcBkdEdn is undefined and MUST be 
						  //ignored
						  //Office Word 2003 emits information at offset fcBkdEdn. Neither Word 
						  //97, Word 2000, Word 2002, Office Word 2007, Word 2010, nor Word 2013 
						  //emit this information
						  //Office Word 2003 reads this information. Word 97, Word 2000, Word 
						  //2002, Office Word 2007, Word 2010, and Word 2013 ignore this 
						  //information
	uint32_t lcbBkdEdn;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
						  //the deprecated endnote text flow break cache at offset fcBkdEdn in the
						  //Table Stream
	uint32_t fcAfdEdn;    //(4 bytes):   An unsigned integer that specifies an offset in the Table
						  //Stream. Deprecated endnote author filter cache begins at this offset.
						  //Information SHOULD NOT be emitted at this offset and SHOULD be 
						  //ignored. If lcbAfdEdn is zero, then fcAfdEdn is undefined and MUST be 
						  //ignored
						  //Office Word 2003 emits information at offset fcAfdEdn. Neither Word 
						  //97, Word 2000, Word 2002, Office Word 2007, Word 2010, nor Word 2013 
						  //emit this information
						  //Office Word 2003 reads this information. Word 97, Word 2000, Word 
						  //2002, Office Word 2007, Word 2010, and Word 2013 ignore this 
						  //information
	uint32_t lcbAfdEdn;   //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
						  //the deprecated endnote author filter cache at offset fcAfdEdn in the
						  //Table Stream
	uint32_t fcAfd;       //(4 bytes):   An unsigned integer that specifies an offset in the Table
						  //Stream. A deprecated AFD structure begins at this offset. Information
						  //SHOULD NOT be emitted at this offset and SHOULD be ignored. If lcbAfd
						  //is zero, fcAfd is undefined and MUST be ignored
						  //Office Word 2003 emits information at offset fcAfd. Neither Word 97, 
						  //Word 2000, Word 2002, Office Word 2007, Word 2010, nor Word 2013 emit 
						  //information at this offset
						  //Office Word 2003 reads this information. Word 97, Word 2000, Word 
						  //2002, Office Word 2007, Word 2010, and Word 2013 ignore this 
						  //information
	uint32_t lcbAfd;      //(4 bytes):   An unsigned integer that specifies the size, in bytes, of
						  //the deprecated AFD structure at offset fcAfd in the Table Stream
} FibRgFcLcb2003;

/*
 * FibRgFcLcb2007
 * The FibRgFcLcb2007 structure is a variable-sized portion of the Fib. It extends the 
 * FibRgFcLcb2003.
 */
typedef struct FibRgFcLcb2007 
{
	FibRgFcLcb2003 rgFcLcb2003;//(1312 bytes): The contained FibRgFcLcb2003.
	uint32_t fcPlcfmthd;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbPlcfmthd; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcSttbfBkmkMoveFrom;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbSttbfBkmkMoveFrom; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcPlcfBkfMoveFrom;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbPlcfBkfMoveFrom; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcPlcfBklMoveFrom;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbPlcfBklMoveFrom; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcSttbfBkmkMoveTo;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbSttbfBkmkMoveTo; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcPlcfBkfMoveTo;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbPlcfBkfMoveTo; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcPlcfBklMoveTo;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbPlcfBklMoveTo; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcUnused1;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbUnused1; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcUnused2;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbUnused2; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcUnused3;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbUnused3; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcSttbfBkmkArto;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbSttbfBkmkArto; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcPlcfBkfArto;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbPlcfBkfArto; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcPlcfBklArto;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbPlcfBklArto; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcArtoData;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbArtoData; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcUnused4;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbUnused4; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcUnused5;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbUnused5; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcUnused6;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbUnused6; //(4 bytes):   This value MUST be zero, and MUST be ignored
	uint32_t fcOssTheme;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbOssTheme; //(4 bytes):   This value MUST be zero, and MUST be ignored
						  //Neither Office Word 2007, Word 2010, nor Word 2013 write 0 here, but 
						  //all three ignore this value when loading files
	uint32_t fcColorSchemeMapping;  //(4 bytes):   This value is undefined and MUST be ignored
	uint32_t lcbColorSchemeMapping; //(4 bytes):   This value MUST be zero, and MUST be ignored
						  //Neither Office Word 2007, Word 2010, nor Word 2013 write 0 here, but 
						  //all three ignore this value when loading files

} FibRgFcLcb2007;

/*
 * FibRgCswNew
 * The FibRgCswNew structure is an extension to the Fib structure that exists only if Fib.cswNew 
 * is nonzero.
 */
typedef struct FibRgCswNew 
{
	uint16_t nFibNew;//(2 bytes): An unsigned integer that specifies the version number of the 
					      //file format that is used. This value MUST be one of the following
					      //0x00D9, 0x0101, 0x010C, 0x0112
	uint16_t rgCswNewData[4];
} FibRgCswNew;

enum rgCswNewData_t {
	FibRgCswNewData2000_t,
	FibRgCswNewData2007_t,
};

struct nFibNew2rgCswNewData {
	uint16_t nFibNew;
	enum rgCswNewData_t rgCswNewData;
};

const struct nFibNew2rgCswNewData nFibNew2rgCswNewDataTable[] = {
	{0x00D9, FibRgCswNewData2000_t}, 
	{0x0101, FibRgCswNewData2000_t}, 
	{0x010C, FibRgCswNewData2000_t}, 
	{0x0112, FibRgCswNewData2007_t}, 
};

enum rgCswNewData_t rgCswNewData_get(uint16_t nFibNew){
	if (nFibNew == 0x0112)
		return FibRgCswNewData2007_t;
	return FibRgCswNewData2000_t;
}

/*
 * FibRgCswNewData2000
 * The FibRgCswNewData2000 structure is a variable-sized portion of the Fib.
 */
typedef struct FibRgCswNewData2000 
{
	uint16_t cQuickSavesNew;//(2 bytes): An unsigned integer that specifies the number of times 
					      //that this document was incrementally saved since the last full save. 
						  //This value MUST be between 0 and 0x000F, inclusively 
} FibRgCswNewData2000;

/*
 * FibRgCswNewData2007
 * The FibRgCswNewData2007 structure is a variable-sized portion of the Fib. It extends the 
 * FibRgCswNewData2000.
 */
typedef struct FibRgCswNewData2007 
{
	FibRgCswNewData2000 rgCswNewData2000;//The contained FibRgCswNewData2000.
	uint16_t lidThemeOther;//(2 bytes): This value is undefined and MUST be ignored 
	uint16_t lidThemeFE;   //(2 bytes): This value is undefined and MUST be ignored 
	uint16_t lidThemeCS;   //(2 bytes): This value is undefined and MUST be ignored 
} FibRgCswNewData2007;

/*
 * Sprm
 * The Sprm structure specifies a modification to a property of a character, paragraph, 
 * table, or section.
 */ 
struct Sprm {
	uint16_t ispmdAsgcspra; //ispmd (9 bits): An unsigned integer that, when combined 
							//with fSpec, specifies the property being modified. See 
							//the tables in the Single Property Modifiers section (2.6) 
							//for the complete list of valid ispmd, fSpec, spra 
							//combinations for each sgc. 
							// 
							// A - fSpec (1 bit): When combined with ispmd, specifies the 
							// property being modified. See the tables in the Single 
							// Property Modifiers section (2.6) for the complete list of 
							// valid ispmd, fSpec, spra combinations for each sgc.
							// 
							// sgc (3 bits): An unsigned integer that specifies the kind of
							// document content to which this Sprm applies. The following 
							// table specifies the valid values and their meanings.
							// Sgc: Meaning
							// 1:    Sprm is modifying a paragraph property.
							// 2:    Sprm is modifying a character property.
							// 3:    Sprm is modifying a picture property.
							// 4:    Sprm is modifying a section property.
							// 5:    Sprm is modifying a table property.
							// 
							// spra (3 bits): An unsigned integer that specifies the size 
							// of the operand of this Sprm. The following table specifies 
							// the valid values and their meanings 
							// Spra: Meaning
							// 0 Operand is a ToggleOperand (which is 1 byte in size).
							// 1 Operand is 1 byte.
							// 2 Operand is 2 bytes.
							// 3 Operand is 4 bytes.
							// 4 Operand is 2 bytes.
							// 5 Operand is 2 bytes.
							// 6 Operand is of variable length. The first byte of the 
							// operand indicates the size of the rest of the operand, 
							// except in the cases of sprmTDefTable and sprmPChgTabs.
							// 7 Operand is 3 bytes.
};

uint8_t SprmIspmd(struct Sprm *sprm){
	return sprm->ispmdAsgcspra & 0x1F;
}
uint8_t SprmA(struct Sprm *sprm){
	return (sprm->ispmdAsgcspra & 0x20) >> 9;
}
uint8_t SprmSgc(struct Sprm *sprm){
	return (sprm->ispmdAsgcspra & 0x1C0) >> 10;
}
uint8_t SprmSpra(struct Sprm *sprm){
	return (sprm->ispmdAsgcspra & 0xE00) >> 13;
}


/*
 * Prl
 * The Prl structure is a Sprm that is followed by an operand. The Sprm specifies a property to 
 * modify, and the operand specifies the new value.
 */
struct Prl {
	struct Sprm sprm;   //(2 bytes): Sprm which specifies the property to be modified
	void *operand;      //(variable): The meaning of the operand depends on the sprm(Single 
					    //Property Modifiers).
};

/*
 * PrcData
 * The PrcData structure specifies an array of Prl elements and the size of the array.
 */
struct PrcData {
	uint16_t cbGrpprl; //(2 byte): A signed integer that specifies the size of GrpPrl, in bytes. 
					   //This value MUST be less than or equal to 0x3FA2
	struct Prl *GrpPrl;//(variable):  An array of Prl elements. GrpPrl contains a whole number 
					   //of Prl elements.
};


/*
 * Prc
 * The Prc structure specifies a set of properties for document content that is referenced by a 
 * Pcd structure.
 */
struct Prc {
	uint8_t clxt;         //(1 byte): This value MUST be 0x01
	struct PrcData *data; //(variable):  PrcData that specifies a set of properties.
};

struct FcCompressedSpecialChar {
	uint8_t  byte;
	uint16_t unicodeCharacter;
};

static const struct FcCompressedSpecialChar FcCompressedSpecialChars[] = 
{
	 {0x82, 0x201A},
	 {0x83, 0x0192},
	 {0x84, 0x201E},
	 {0x85, 0x2026},
	 {0x86, 0x2020},
	 {0x87, 0x2021},
	 {0x88, 0x02C6},
	 {0x89, 0x2030},
	 {0x8A, 0x0160},
	 {0x8B, 0x2039},
	 {0x8C, 0x0152},
	 {0x91, 0x2018},
	 {0x92, 0x2019},
	 {0x93, 0x201C},
	 {0x94, 0x201D},
	 {0x95, 0x2022},
	 {0x96, 0x2013},
	 {0x97, 0x2014},
	 {0x98, 0x02DC},
	 {0x99, 0x2122},
	 {0x9A, 0x0161},
	 {0x9B, 0x203A},
	 {0x9C, 0x0153},
	 {0x9F, 0x0178},
};

static int FcCompressedSpecialChar_compare(const void *key, const void *value) {
    const struct FcCompressedSpecialChar *cp1 = key;
    const struct FcCompressedSpecialChar *cp2 = value;
    return cp1->byte - cp2->byte;
}

uint16_t FcCompressedSpecialChar_get(uint16_t nFib){
    struct FcCompressedSpecialChar *result = bsearch(&nFib, FcCompressedSpecialChars,
            sizeof(FcCompressedSpecialChars)/sizeof(FcCompressedSpecialChars[0]),
            sizeof(FcCompressedSpecialChars[0]), FcCompressedSpecialChar_compare);
	if (result)
		return result->unicodeCharacter;
	return 0;
}

/*
 * FcCompressed
 * The FcCompressed structure specifies the location of text in the WordDocument Stream.
 */
struct FcCompressed {
	uint32_t fc; //fc (30 bits): An unsigned integer that specifies an offset in the 
				 //WordDocument Stream where the text starts. If fCompressed is zero, the text 
				 //is an array of 16-bit Unicode characters starting at offset fc. If 
				 //fCompressed is 1, the text starts at offset fc/2 and is an array of 8-bit 
				 //Unicode characters, except for the values which are mapped to Unicode 
				 //characters as follows
				 //0x82 0x201A 
				 //0x83 0x0192 
				 //0x84 0x201E 
				 //0x85 0x2026 
				 //0x86 0x2020 
				 //0x87 0x2021 
				 //0x88 0x02C6 
				 //0x89 0x2030 
				 //0x8A 0x0160 
				 //0x8B 0x2039 
				 //0x8C 0x0152 
				 //0x91 0x2018 
				 //0x92 0x2019 
				 //0x93 0x201C 
				 //0x94 0x201D 
				 //0x95 0x2022 
				 //0x96 0x2013 
				 //0x97 0x2014 
				 //0x98 0x02DC 
				 //0x99 0x2122 
				 //0x9A 0x0161 
				 //0x9B 0x203A 
				 //0x9C 0x0153 
				 //0x9F 0x0178
				 //A - fCompressed (1 bit): A bit that specifies whether the text is compressed.
				 //B - r1 (1 bit): This bit MUST be zero, and MUST be ignored.
};
bool FcCompressed(struct FcCompressed fc){
	///* TODO: byte order */
	if ((fc.fc & 0x40000000) == 0x40000000) //if compressed - then ANSI
		return true;
	return false;
};
uint32_t FcValue(struct FcCompressed fc){
	///* TODO: byte order */
	return fc.fc & 0x3FFFFFFF;	
}

/*
 * Pcd
 * The Pcd structure specifies the location of text in the WordDocument Stream and additional 
 * properties for this text. A Pcd structure is an element of a PlcPcd structure.
 */
struct Pcd {
	uint16_t ABCfR2; //A - fNoParaLast (1 bit): If this bit is 1, the text MUST NOT contain a 
					 //paragraph mark.
					 //B - fR1 (1 bit): This field is undefined and MUST be ignored
					 //C - fDirty (1 bit): This field MUST be 0
					 //fR2 (13 bits): This field is undefined and MUST be ignored
	struct FcCompressed fc;//(4 bytes): An FcCompressed structure that specifies the location of 
				     //the text in the WordDocument Stream 
	uint16_t prm;    //A Prm structure that specifies additional properties for this text.
};

/*
 * PlcPcd
 * The PlcPcd structure is a PLC whose data elements are Pcds (8 bytes each). A PlcPcd MUST NOT 
 * contain duplicate CPs.
 */
struct PlcPcd {
	uint32_t *aCp; //(variable): An array of CPs that specifies the starting points of text 
				   //ranges. The end of each range is the beginning of the next range. All CPs 
				   //MUST be greater than or equal to zero. If any of the fields ccpFtn, ccpHdd, 
				   //ccpAtn, ccpEdn, ccpTxbx, or ccpHdrTxbx from FibRgLw97 are nonzero, then the 
				   //last CP MUST be equal to the sum of those fields plus ccpText+1. Otherwise, 
				   //the last CP MUST be equal to ccpText.
	uint32_t aCPl; //number of CP in array
	struct Pcd *aPcd;//(variable):  An array of Pcds (8 bytes each) that specify the location of 
				   //text in the WordDocument stream and any additional properties of the text. 
				   //If aPcd[i].fc.fCompressed is 1, then the byte offset of the last character 
				   //of the text referenced by aPcd[i] is given by the following.
				   //(aPcd[i].fc.fc/2) + aCP[i+1] - aCP[i] - 1;
				   //Otherwise, the byte offset of the last character of the text referenced by 
				   //aPcd[i] is given by the following
				   //aPcd[i].fc.fc + 2(aCP[i+1]  - aCP[i] - 1)
				   //Because aCP MUST be sorted in ascending order and MUST NOT contain 
				   //duplicate CPs, (aCP[i+1]-aCP[i])>0, for all valid indexes i of aPcd. 
				   //Because a PLC MUST contain one more CP than a data element, i+1 is a valid 
				   //index of aCP if i is a valid index of aPcd.
	uint32_t aPcdl;//number of Pcd in array
};

/*
 * Pcdt
 * The Pcdt structure contains a PlcPcd structure and specifies its size.
 */
struct Pcdt {
	uint8_t clxt;    //(1 byte): This value MUST be 0x02
	uint32_t lcb;    //(4 bytes): An unsigned integer that specifies the size, in bytes, of the 
				     //PlcPcd structure.
	struct PlcPcd PlcPcd; //(variable): A PlcPcd structure. As with all Plc elements, the size 
						   //that is specified by lcb MUST result in a whole number of Pcd 
						   //structures in this PlcPcd structure.
};

/*
 * Clx
 * The Clx structure is an array of zero, 1, or more Prcs followed by a Pcdt.
 */
struct Clx {
	struct Prc *RgPrc; //(variable): An array of Prc. If this array is empty, the first byte of
					   //the Clx MUST be 0x02. 0x02 is invalid as the first byte of a Prc, 
					   //but required for the Pcdt.
	struct Pcdt *Pcdt; //(variable): A Pcdt.
};


/*
 * FIB
 * The Fib structure is located at offset 0 of the WordDocument Stream.
 */

typedef struct Fib 
{
	FibBase *base;        //MUST be present and has fixed size
	uint16_t csw;         //(2 bytes): An unsigned integer that specifies the count of 16-bit 
						  //values corresponding to fibRgW that follow. MUST be 0x000E. 
	FibRgW97 *rgW97;      //Fib.csw * 2 bytes
    uint16_t cslw;        //(2 bytes): An unsigned integer that specifies the count of 32-bit 
						  //values corresponding to fibRgLw that follow. MUST be 0x0016. 
	FibRgLw97 *rgLw97;    //Fib.cslw * 4 bytes
	uint16_t cbRgFcLcb;   //(2 bytes): An unsigned integer that specifies the count of 64-bit 
						  //values corresponding to fibRgFcLcbBlob that follow. This MUST be one 
						  //of the following values, depending on the value of nFib.
	uint32_t *rgFcLcb;    //Fib.cbRgFcLcb * 8 bytes
	uint16_t cswNew;      //(2 bytes): An unsigned integer that specifies the count of 16-bit 
						  //values corresponding to fibRgCswNew that follow. This MUST be one of 
						  //the following values, depending on the value of nFib. 
	FibRgCswNew *rgCswNew;
} Fib;

/*
 * MS-DOC Structure.
 */

typedef struct cfb_doc 
{
	FILE *WordDocument;   //document stream
	FILE *Table;          //table stream
	
	Fib  fib;             //File information block
	struct Clx clx;       //clx data
	bool byteOrder;				//need to change byte order
} cfb_doc_t;


/*
 * How to read the FIB
 * The Fib structure is located at offset 0 of the WordDocument Stream. Given the variable size of
 * the Fib, the proper way to load it is the following:
 * 1.  Set all bytes of the in-memory version of the Fib being used to 0. It is recommended to use 
 *     the largest version of the Fib structure as the in-memory version.
 * 2.  Read the entire FibBase, which MUST be present and has fixed size.
 * 3.  Read Fib.csw.
 * 4.  Read the minimum of Fib.csw * 2 bytes and the size, in bytes, of the in-memory version of 
 *     FibRgW97 into FibRgW97.
 * 5.  If the application expects fewer bytes than indicated by Fib.csw, advance by the difference
 *     thereby skipping the unknown portion of FibRgW97.
 * 6.  Read Fib.cslw.
 * 7.  Read the minimum of Fib.cslw * 4 bytes and the size, in bytes, of the in-memory version of
 *     FibRgLw97 into FibRgLw97.
 * 8.  If the application expects fewer bytes than indicated by Fib.cslw, advance by the 
 *     difference thereby skipping the unknown portion of FibRgLw97.
 * 9.  Read Fib.cbRgFcLcb.
 * 10. Read the minimum of Fib.cbRgFcLcb * 8 bytes and the size, in bytes, of the in-memory 
 *     version of FibRgFcLcb into FibRgFcLcb.
 * 11. If the application expects fewer bytes than indicated by Fib.cbRgFcLcb, advance by the 
 *     difference, thereby skipping the unknown portion of FibRgFcLcb.
 * 12. Read Fib.cswNew.
 * 13. Read the minimum of Fib.cswNew * 2 bytes and the size, in bytes, of the in-memory version 
 *     of FibRgCswNew into FibRgCswNew.
*/

int _cfb_doc_fib_init(Fib *fib, FILE *fp, struct cfb *cfb){
#ifdef DEBUG
	LOG("start _cfb_doc_fib_init\n");
#endif

	fib->base = NULL;
	fib->csw = 0;
	fib->rgW97 = NULL;
	fib->cslw = 0;
	fib->rgLw97 = NULL;
	fib->cbRgFcLcb = 0;
	fib->rgFcLcb = NULL;
	fib->cswNew = 0;
	fib->rgCswNew = NULL;

	//allocate fibbase
#ifdef DEBUG
	LOG("_cfb_doc_fib_init: allocate fibbase\n");
#endif
	
	fib->base = (FibBase *)malloc(32);
	if (!fib->base)
		return DOC_ERR_ALLOC;

	//read fibbase
#ifdef DEBUG
	LOG("_cfb_doc_fib_init: read fibbase\n");
#endif
	
	if (fread(fib->base, 32, 1, fp) != 1){
		free(fib->base);
		return DOC_ERR_FILE;
	}
	if (cfb->biteOrder){
		fib->base->wIdent        = bo_16_sw(fib->base->wIdent);
		fib->base->nFib          = bo_16_sw(fib->base->nFib);
		fib->base->lid           = bo_16_sw(fib->base->lid);
		fib->base->pnNext        = bo_16_sw(fib->base->pnNext);
		fib->base->ABCDEFGHIJKLM = bo_16_sw(fib->base->ABCDEFGHIJKLM);
		fib->base->nFibBack      = bo_16_sw(fib->base->nFibBack);
		fib->base->lKey          = bo_32_sw(fib->base->lKey);
	}
	
	//check wIdent
#ifdef DEBUG
	LOG("_cfb_doc_fib_init: check wIdent: 0x%x\n", fib->base->wIdent);
#endif	
	if (fib->base->wIdent != 0xA5EC){
		free(fib->base);
		return DOC_ERR_HEADER;
	}	

#ifdef DEBUG
	LOG("_cfb_doc_fib_init: read csw\n");
#endif	
	//read Fib.csw
	if (fread(&(fib->csw), 2, 1, fp) != 1){
		free(fib->base);
		return DOC_ERR_FILE;
	}
	if (cfb->biteOrder){
		fib->csw = bo_16_sw(fib->csw);
	}

	//check csw
#ifdef DEBUG
	LOG("_cfb_doc_fib_init: check csw: 0x%x\n", fib->csw);
#endif		
	if (fib->csw != 14) {
		free(fib->base);
		return DOC_ERR_HEADER;
	}

#ifdef DEBUG
	LOG("_cfb_doc_fib_init: allocate FibRgW97\n");
#endif	
	//allocate FibRgW97
	fib->rgW97 = (FibRgW97 *)malloc(28);
	if (!fib->rgW97){
		free(fib->base);
		return DOC_ERR_ALLOC;
	}

	//read FibRgW97
#ifdef DEBUG
	LOG("_cfb_doc_fib_init: read FibRgW97\n");
#endif
	if (fread(fib->rgW97, 28, 1, fp) != 1){
		free(fib->base);
		free(fib->rgW97);
		return DOC_ERR_FILE;
	}
	if (cfb->biteOrder){
		fib->rgW97->lidFE = bo_16_sw(fib->rgW97->lidFE);
	}

#ifdef DEBUG
	LOG("_cfb_doc_fib_init: read Fib.cslw\n");
#endif	
	//read Fib.cslw
	if (fread(&(fib->cslw), 2, 1, fp) != 1){
		free(fib->base);
		free(fib->rgW97);
		return DOC_ERR_FILE;
	}
	if (cfb->biteOrder){
		fib->cslw = bo_16_sw(fib->cslw);
	}

#ifdef DEBUG
	LOG("_cfb_doc_fib_init: check cslw: 0x%x\n", fib->cslw);
#endif	
	//check cslw
	if (fib->cslw != 22) {
		free(fib->base);
		free(fib->rgW97);
		return DOC_ERR_HEADER;
	}	

#ifdef DEBUG
	LOG("_cfb_doc_fib_init: allocate FibRgLw97\n");
#endif	
	//allocate FibRgLw97
	fib->rgLw97 = (FibRgLw97 *)malloc(88);
	if (!fib->rgLw97){
		free(fib->base);
		free(fib->rgW97);
		return DOC_ERR_ALLOC;
	}
	
#ifdef DEBUG
	LOG("_cfb_doc_fib_init: read Fib.FibRgLw97\n");
#endif	
	//read FibRgLw97
	if (fread(fib->rgLw97, 88, 1, fp) != 1){
		free(fib->base);
		free(fib->rgW97);
		free(fib->rgLw97);
		return DOC_ERR_FILE;
	}	
	if (cfb->biteOrder){
		fib->rgLw97->cbMac      = bo_32_sw(fib->rgLw97->cbMac);
		fib->rgLw97->ccpText    = bo_32_sw(fib->rgLw97->ccpText);
		fib->rgLw97->ccpFtn     = bo_32_sw(fib->rgLw97->ccpFtn);
		fib->rgLw97->ccpHdd     = bo_32_sw(fib->rgLw97->ccpHdd);
		fib->rgLw97->ccpAtn     = bo_32_sw(fib->rgLw97->ccpAtn);
		fib->rgLw97->ccpEdn     = bo_32_sw(fib->rgLw97->ccpEdn);
		fib->rgLw97->ccpTxbx    = bo_32_sw(fib->rgLw97->ccpTxbx);
		fib->rgLw97->ccpHdrTxbx = bo_32_sw(fib->rgLw97->ccpHdrTxbx);
	}
	
#ifdef DEBUG
	LOG("_cfb_doc_fib_init: read Fib.cbRgFcLcb\n");
#endif	
	//read Fib.cbRgFcLcb
	if (fread(&(fib->cbRgFcLcb), 2, 1, fp) != 1){
		free(fib->base);
		free(fib->rgW97);
		free(fib->rgLw97);
		return DOC_ERR_FILE;
	}
	if (cfb->biteOrder){
		fib->cbRgFcLcb = bo_16_sw(fib->cbRgFcLcb);
	}
	
#ifdef DEBUG
	LOG("cbRgFcLcb: 0x%x\n", fib->cbRgFcLcb);
#endif	

#ifdef DEBUG
	LOG("_cfb_doc_fib_init: allocate FibRgLw97 with size: %d\n", fib->cbRgFcLcb*8);
#endif	
	//allocate rgFcLcb
	fib->rgFcLcb = (uint32_t *)malloc(fib->cbRgFcLcb*8);
	if (!fib->rgFcLcb){
		free(fib->base);
		free(fib->rgW97);
		free(fib->rgLw97);
		return DOC_ERR_ALLOC;
	}	

#ifdef DEBUG
	LOG("_cfb_doc_fib_init: read Fib.rgFcLcb\n");
#endif	
	//read rgFcLcb
	if (fread(fib->rgFcLcb, 8, fib->cbRgFcLcb, fp) != fib->cbRgFcLcb){
		free(fib->base);
		free(fib->rgW97);
		free(fib->rgLw97);
		free(fib->rgFcLcb);
		return DOC_ERR_FILE;
	}	
	//if (cfb->biteOrder){
		//int i;
		//for (i = 0; i < fib->cbRgFcLcb/4; ++i) {
			//fib->rgFcLcb[i] = bo_32_sw(fib->rgFcLcb[i]);	
		//}
	//}

#ifdef DEBUG
	LOG("_cfb_doc_fib_init: read Fib.cswNew\n");
#endif	
	//read Fib.cswNew
	fread(&(fib->cswNew), 2, 1, fp);

#ifdef DEBUG
	LOG("cswNew: 0x%x\n", fib->cswNew);
#endif	
	if (cfb->biteOrder){
		fib->cswNew = bo_16_sw(fib->cswNew);
	}

	if (fib->cswNew > 0){
#ifdef DEBUG
	LOG("_cfb_doc_fib_init: allocate FibRgCswNew with size: %d\n", fib->cswNew * 2);
#endif		
		//allocate FibRgCswNew
		fib->rgCswNew = (FibRgCswNew *)malloc(fib->cswNew * 2);
		if (!fib->rgFcLcb){
			free(fib->base);
			free(fib->rgW97);
			free(fib->rgLw97);
			free(fib->rgFcLcb);
			return DOC_ERR_ALLOC;
		}	

#ifdef DEBUG
	LOG("_cfb_doc_fib_init: read FibRgCswNew\n");
#endif		
		//read FibRgCswNew
		if (fread(fib->rgCswNew, 2, fib->cswNew, fp) != fib->cswNew){
			free(fib->base);
			free(fib->rgW97);
			free(fib->rgLw97);
			free(fib->rgFcLcb);
			return DOC_ERR_FILE;
		}	
		if (cfb->biteOrder){
			fib->rgCswNew->nFibNew = bo_16_sw(fib->rgCswNew->nFibNew);
			int i;
			for (i = 0; i < 4; ++i) {
				fib->rgCswNew->rgCswNewData[i] = bo_16_sw(fib->rgCswNew->rgCswNewData[i]);
			}
		}
	}
#ifdef DEBUG
	LOG("_cfb_doc_fib_init done\n");
#endif	
	return 0;
};

FILE *_table_stream(cfb_doc_t *doc, struct cfb *cfb){
	char *table = "0Table";
	Fib *fib = &doc[0].fib;
	if (FibBaseG(fib[0].base))
		table = "1Table";
#ifdef DEBUG
	LOG("_table_stream: table: %s\n", table);
#endif	
	return cfb_get_stream(cfb, table);
}

int _plcpcd_init(struct PlcPcd * PlcPcd, uint32_t len, cfb_doc_t *doc){
#ifdef DEBUG
	LOG("start _plcpcd_init\n");
#endif	
	
	int i;

	//get lastCP
	uint32_t lastCp = 
			doc->fib.rgLw97->ccpFtn +
			doc->fib.rgLw97->ccpHdd +
			doc->fib.rgLw97->reserved3 + //Mcr
			doc->fib.rgLw97->ccpAtn +
			doc->fib.rgLw97->ccpEdn +
			doc->fib.rgLw97->ccpTxbx +
			doc->fib.rgLw97->ccpHdrTxbx;
	
	if (lastCp)
		lastCp += 1 + doc->fib.rgLw97->ccpText;
	else
		lastCp = doc->fib.rgLw97->ccpText;

#ifdef DEBUG
	LOG("_plcpcd_init: lastCp: %d\n", lastCp);
#endif	

#ifdef DEBUG
	LOG("_plcpcd_init: allocate aCP\n");
#endif	
	//allocate aCP
	PlcPcd->aCp = malloc(4);
	if (!PlcPcd->aCp){
		free(PlcPcd);	
		return -1;
	}

	//read aCP
	i=0;
	uint32_t ch;
	while(fread(&ch, 4, 1, doc->Table) == 1){
		if (doc->byteOrder){
			ch = bo_32_sw(ch);
		}
		PlcPcd->aCp[i++] = ch;
#ifdef DEBUG
	LOG("_plcpcd_init: aCp[%d]: %u\n", PlcPcd->aCp[i], ch);
#endif		
		if (ch == lastCp)
			break;

		//realloc aCp
#ifdef DEBUG
	LOG("_plcpcd_init: realloc aCP with size: %u\n", (i+1)*4);
#endif
		void *ptr = realloc(PlcPcd->aCp, (i+1)*4);
		if(!ptr)
			break;
		PlcPcd->aCp = ptr;
	}
#ifdef DEBUG
	LOG("_plcpcd_init: number of cp in array: %d\n", i);
#endif	
	//number of cp in array
	PlcPcd->aCPl = i;

	//read PCD - has 64bit
	uint32_t size = len - i*4;
#ifdef DEBUG
	LOG("_plcpcd_init: allocate aCP with size: %u\n", size);
#endif	
	PlcPcd->aPcd = malloc(size);
	if (!PlcPcd->aPcd){
		free(PlcPcd->aCp);	
		free(PlcPcd);	
		return -1;
	}
	fread(PlcPcd->aPcd, size, 1, doc->Table);
	//number of Pcd in array
	PlcPcd->aPcdl = size / 8;
#ifdef DEBUG
	LOG("_plcpcd_init: number of Pcd in array: %d\n", PlcPcd->aPcdl);
#endif	
	
	if (doc->byteOrder){
		for (i = 0; i < PlcPcd->aPcdl; ++i) {
			PlcPcd->aPcd[i].ABCfR2 = bo_16_sw(PlcPcd->aPcd[i].ABCfR2); 
			PlcPcd->aPcd[i].prm = bo_16_sw(PlcPcd->aPcd[i].prm); 
			PlcPcd->aPcd[i].fc.fc = bo_32_sw(PlcPcd->aPcd[i].fc.fc); 
		}
	}

#ifdef DEBUG
	LOG("_plcpcd_init: PlcPcd->aPcd[%d]: ABCfR2: 0x%x, FC: 0x%x, PRM: 0x%x\n", i, PlcPcd->aPcd[i].ABCfR2, PlcPcd->aPcd[i].fc.fc, PlcPcd->aPcd[i].prm);
#endif	
	
#ifdef DEBUG
	LOG("_plcpcd_init done\n");
#endif	

	return 0;
}

int _clx_init(struct Clx *clx, uint32_t fcClx, uint32_t lcbClx, cfb_doc_t *doc){

#ifdef DEBUG
	LOG("start _clx_init\n");
#endif

	//get clx
	uint8_t ch;
	fseek(doc->Table, fcClx, SEEK_SET);
	fread(&ch, 1, 1, doc->Table);
#ifdef DEBUG
	LOG("_clx_init: first bite of CLX: 0x%x\n", ch);
#endif

	if (ch == 0x01){ //we have RgPrc (Prc array)
#ifdef DEBUG
	LOG("_clx_init: we have RgPrc (Prc array)\n");
#endif		
		//allocate RgPrc
#ifdef DEBUG
	LOG("_clx_init: allocate RgPrc\n");
#endif
		clx->RgPrc = malloc(sizeof(struct Prc));
		if (!clx->RgPrc)
			return DOC_ERR_ALLOC;
		
		int16_t cbGrpprl; //the first 2 bite of PrcData - signed integer
		fread(&cbGrpprl, 2, 1, doc->Table);
		if (doc->byteOrder){
			cbGrpprl = bo_16_sw(cbGrpprl);
		}
#ifdef DEBUG
	LOG("_clx_init: the first 2 bite of PrcData is cbGrpprl: 0x%x\n", cbGrpprl);
#endif		
		if (cbGrpprl > 0x3FA2) //error
			return DOC_ERR_FILE;		
		
		//allocate RgPrc->data 
#ifdef DEBUG
	LOG("_clx_init: allocate RgPrc->data\n");
#endif
		clx->RgPrc->data = malloc(sizeof(struct PrcData));
		if (!clx->RgPrc->data)
			return DOC_ERR_ALLOC;

		clx->RgPrc->data->cbGrpprl = cbGrpprl;

		//allocate GrpPrl
#ifdef DEBUG
	LOG("_clx_init: allocate GrpPrl with size: 0x%x\n", cbGrpprl);
#endif		
		clx->RgPrc->data->GrpPrl = malloc(cbGrpprl);
		if (!clx->RgPrc->data->GrpPrl)
			return DOC_ERR_ALLOC;
		
		//read GrpPrl
#ifdef DEBUG
	LOG("_clx_init: read GrpPrl\n");
#endif		
		fread(clx->RgPrc->data->GrpPrl, cbGrpprl, 1, doc->Table);
		/* TODO:  parse GrpPrl + byteOrder */

		//read ch again
#ifdef DEBUG
	LOG("_clx_init: again first bite of CLX: 0x%x\n", ch);
#endif		
		fread(&ch, 1, 1, doc->Table);
	}	

	//get PlcPcd
#ifdef DEBUG
	LOG("_clx_init: allocate PlcPcd\n");
#endif
	clx->Pcdt = malloc(sizeof(struct Pcdt));
	if (!clx->Pcdt)
		return DOC_ERR_ALLOC;	

	//read Pcdt->clxt - this must be 0x02
	clx->Pcdt->clxt = ch;
#ifdef DEBUG
	LOG("_clx_init: Pcdt->clxt: 0x%x\n", clx->Pcdt->clxt);
#endif	
	if (clx->Pcdt->clxt != 0x02) { //some error
		return DOC_ERR_FILE;		
	}

	//read lcb;
	fread(&(clx->Pcdt->lcb), 4, 1, doc->Table);	
	if (doc->byteOrder){
		clx->Pcdt->lcb = bo_32_sw(clx->Pcdt->lcb);
	}
#ifdef DEBUG
	LOG("_clx_init: Pcdt->lcb: %d\n", clx->Pcdt->lcb);
#endif	

	//get PlcPcd
	_plcpcd_init(&(clx->Pcdt->PlcPcd), clx->Pcdt->lcb, doc);
	
#ifdef DEBUG
	LOG("_clx_init: aCP: %d, PCD: %d\n", clx->Pcdt->PlcPcd.aCPl, clx->Pcdt->PlcPcd.aPcdl);
#endif		

#ifdef DEBUG
	LOG("_clx_init done\n");
#endif
	
	return 0;
}

int cfb_doc_init(cfb_doc_t *doc, struct cfb *cfb){
#ifdef DEBUG
	LOG("start cfb_doc_init\n");
#endif
	
	int ret = 0;
	//get byte order
	doc->byteOrder = cfb->biteOrder;
	
	//get WordDocument
	FILE *fp = cfb_get_stream(cfb, "WordDocument");
	if (!fp)	
		return DOC_ERR_FILE;
	fseek(fp, 0, SEEK_SET);
	doc->WordDocument = fp;

	//init FIB
	_cfb_doc_fib_init(&(doc->fib), doc->WordDocument, cfb);

	//get table
	doc->Table = _table_stream(doc, cfb);
	if (!doc->Table){
		//printf("Can't get Table stream\n"); 
		return DOC_ERR_FILE;
	}

	//get CLX
	//All versions of the FIB contain exactly one FibRgFcLcb97 
	FibRgFcLcb97 *rgFcLcb97 = (FibRgFcLcb97 *)(doc->fib.rgFcLcb);
	//FibRgFcLcb97.fcClx specifies the offset in the Table Stream of a Clx
	uint32_t fcClx = rgFcLcb97->fcClx;
	if (cfb->biteOrder)
		fcClx = bo_32_sw(fcClx);
#ifdef DEBUG
	LOG("fcClx: %d\n", fcClx);
#endif
	//FibRgFcLcb97.lcbClx specifies the size, in bytes, of that Clx
	uint32_t lcbClx = rgFcLcb97->lcbClx;
	if (cfb->biteOrder)
		lcbClx = bo_32_sw(lcbClx);
#ifdef DEBUG
	LOG("lcbClx: %d\n", lcbClx);
#endif	

	//Read the Clx from the Table Stream
	ret = _clx_init(&(doc->clx), rgFcLcb97->fcClx, rgFcLcb97->lcbClx, doc);
	if (ret)
		return ret;	

#ifdef DEBUG
	LOG("cfb_doc_init done\n");
#endif	
	return 0;
}
void _get_text(cfb_doc_t *doc, struct PlcPcd *PlcPcd,
		void *user_data,
		int (*text)(
			void *user_data,
			char *str
			)		
		){
	// get char for each CP
	// CPs are in range in aCp
	int i; //aCp iterator
	uint32_t cp = 0; //CP (char position)
	for (i=0; i < PlcPcd->aPcdl; i++){

/*
 * The Clx contains a Pcdt, and the Pcdt contains a PlcPcd. Find the largest i such that 
 * PlcPcd.aCp[i] ≤ cp. As with all Plcs, the elements of PlcPcd.aCp are sorted in ascending order.
 * Recall from the definition of a Plc that the aCp array has one more element than the aPcd array.
 * Thus, if the last element of PlcPcd.aCp is less than or equal to cp, cp is outside the range of
 * valid character positions in this document
 */
		
/*
 * PlcPcd.aPcd[i] is a Pcd. Pcd.fc is an FcCompressed that specifies the location in the 
 * WordDocument Stream of the text at character position PlcPcd.aCp[i].
 */
		struct FcCompressed fc = PlcPcd->aPcd[i].fc;	
		if (FcCompressed(fc)){
/*
 * If FcCompressed.fCompressed is 1, the character at position cp is an 8-bit ANSI character at 
 * offset (FcCompressed.fc / 2) + (cp - PlcPcd.aCp[i]) in the WordDocument Stream, unless it is 
 * one of the special values in the table defined in the description of FcCompressed.fc. This is 
 * to say that the text at character position PlcPcd.aCP[i] begins at offset FcCompressed.fc / 2 
 * in the WordDocument Stream and each character occupies one byte.
 */			
			//ANSI
			DWORD off = (FcValue(fc) / 2) + PlcPcd->aCp[i];
			fseek(doc->WordDocument, off, SEEK_SET);	
			for (cp = PlcPcd->aCp[i]; cp < PlcPcd->aCp[i+1]; cp++){
				//DWORD off = (FcValue(fc) / 2) + (cp - PlcPcd->aCp[i]);
				//fseek(doc->WordDocument, off, SEEK_SET);	
				char c[2] = {0};
				fread(c, 1, 1, doc->WordDocument);			
				text(user_data, c);
			}

		} else {
/*
 * If FcCompressed.fCompressed is zero, the character at position cp is a 16-bit Unicode character
 * at offset FcCompressed.fc + 2(cp - PlcPcd.aCp[i]) in the WordDocument Stream. This is to say
 * that the text at character position PlcPcd.aCP[i] begins at offset FcCompressed.fc in the 
 * WordDocument Stream and each character occupies two bytes.
 */			
			//UNICODE 16
			DWORD off = FcValue(fc) + 2*PlcPcd->aCp[i];
			fseek(doc->WordDocument, off, SEEK_SET);	
			for (cp = PlcPcd->aCp[i]; cp < PlcPcd->aCp[i+1]; cp++){
				//DWORD off = FcValue(fc) + 2*(cp - PlcPcd->aCp[i]);
				//fseek(doc->WordDocument, off, SEEK_SET);	
				WORD u;
				fread(&u, 2, 1, doc->WordDocument);
				if (doc->byteOrder){
					u = bo_16_sw(u);
				}
				char utf8[4]={0};
				_utf16_to_utf8(&u, 1, utf8);
				text(user_data, utf8);
			}
		}
	}
}

static void _get_text_for_cp(cfb_doc_t *doc, struct PlcPcd *PlcPcd, uint32_t cp, uint32_t len,
		void *user_data,
		int (*text)(
			void *user_data,
			char *str
			)		
		){
	
	int i, l=0; //iterator
	while (l++ < len){

/*
 * The Clx contains a Pcdt, and the Pcdt contains a PlcPcd. Find the largest i such that 
 * PlcPcd.aCp[i] ≤ cp. As with all Plcs, the elements of PlcPcd.aCp are sorted in ascending order.
 * Recall from the definition of a Plc that the aCp array has one more element than the aPcd array.
 * Thus, if the last element of PlcPcd.aCp is less than or equal to cp, cp is outside the range of
 * valid character positions in this document
 */
		for (i = 0; i < PlcPcd->aPcdl; ++i) {
			if (PlcPcd->aCp[i] > cp) {
				--i;
				break;
			}	
		}
		
/*
 * PlcPcd.aPcd[i] is a Pcd. Pcd.fc is an FcCompressed that specifies the location in the 
 * WordDocument Stream of the text at character position PlcPcd.aCp[i].
 */
		struct FcCompressed fc = PlcPcd->aPcd[i].fc;	
		if (FcCompressed(fc)){
/*
 * If FcCompressed.fCompressed is 1, the character at position cp is an 8-bit ANSI character at 
 * offset (FcCompressed.fc / 2) + (cp - PlcPcd.aCp[i]) in the WordDocument Stream, unless it is 
 * one of the special values in the table defined in the description of FcCompressed.fc. This is 
 * to say that the text at character position PlcPcd.aCP[i] begins at offset FcCompressed.fc / 2 
 * in the WordDocument Stream and each character occupies one byte.
 */			
			//ANSI
			DWORD off = (FcValue(fc) / 2) + (cp - PlcPcd->aCp[i]);
			fseek(doc->WordDocument, off, SEEK_SET);	
			char c[2] = {0};
			fread(c, 1, 1, doc->WordDocument);			
			text(user_data, c);

		} else {
/*
 * If FcCompressed.fCompressed is zero, the character at position cp is a 16-bit Unicode character
 * at offset FcCompressed.fc + 2(cp - PlcPcd.aCp[i]) in the WordDocument Stream. This is to say
 * that the text at character position PlcPcd.aCP[i] begins at offset FcCompressed.fc in the 
 * WordDocument Stream and each character occupies two bytes.
 */			
			//UNICODE 16
			DWORD off = FcValue(fc) + 2*(cp - PlcPcd->aCp[i]);
			fseek(doc->WordDocument, off, SEEK_SET);	
			WORD u;
			fread(&u, 2, 1, doc->WordDocument);
			if (doc->byteOrder){
				u = bo_16_sw(u);
			}
			char utf8[4]={0};
			_utf16_to_utf8(&u, 1, utf8);
			text(user_data, utf8);
		}

		// iterate
		cp++;
	}
}

/*
 * Retrieving Text
 * The following algorithm specifies how to find the text at a particular character position 
 * (cp). Negative character positions are not valid.
 * 1. Read the FIB from offset zero in the WordDocument Stream.
 * 2. All versions of the FIB contain exactly one FibRgFcLcb97, though it can be nested in a 
 * larger structure. FibRgFcLcb97.fcClx specifies the offset in the Table Stream of a Clx. 
 * FibRgFcLcb97.lcbClx specifies the size, in bytes, of that Clx. Read the Clx from the Table 
 * Stream.
 * 3. The Clx contains a Pcdt, and the Pcdt contains a PlcPcd. Find the largest i such that 
 * PlcPcd.aCp[i] ≤ cp. As with all Plcs, the elements of PlcPcd.aCp are sorted in ascending 
 * order. Recall from the definition of a Plc that the aCp array has one more element than the 
 * aPcd array. Thus, if the last element of PlcPcd.aCp is less than or equal to cp, cp is 
 * outside the range of valid character positions in this document.
 * 4. PlcPcd.aPcd[i] is a Pcd. Pcd.fc is an FcCompressed that specifies the location in the 
 * WordDocument Stream of the text at character position PlcPcd.aCp[i].
 * 5. If FcCompressed.fCompressed is zero, the character at position cp is a 16-bit Unicode 
 * character at offset FcCompressed.fc + 2(cp - PlcPcd.aCp[i]) in the WordDocument Stream. This 
 * is to say that the text at character position PlcPcd.aCP[i] begins at offset FcCompressed.fc 
 * in the WordDocument Stream and each character occupies two bytes.
 * 6. If FcCompressed.fCompressed is 1, the character at position cp is an 8-bit ANSI character 
 * at offset (FcCompressed.fc / 2) + (cp - PlcPcd.aCp[i]) in the WordDocument Stream, unless it 
 * is one of the special values in the table defined in the description of FcCompressed.fc. This 
 * is to say that the text at character position PlcPcd.aCP[i] begins at offset FcCompressed.
 * fc / 2 in the WordDocument Stream and each character occupies one byte.
 *
 * Determining Paragraph Boundaries
 * This section specifies how to find the beginning and end character positions of the paragraph 
 * that contains a given character position. The character at the end character position of a 
 * paragraph MUST be a paragraph mark, an end-of-section character, a cell mark, or a TTP mark 
 * (See Overview of Tables). Negative character positions are not valid.
 * To find the character position of the first character in the paragraph that contains a given 
 * character position cp:
 * 1. Follow the algorithm from Retrieving Text up to and including step 3 to find i. Also 
 * remember the FibRgFcLcb97 and PlcPcd found in step 1 of Retrieving Text. If the algorithm 
 * from Retrieving Text specifies that cp is invalid, leave the algorithm.
 * 2. Let pcd be PlcPcd.aPcd[i].
 * 3. Let fcPcd be Pcd.fc.fc. Let fc be fcPcd + 2(cp – PlcPcd.aCp[i]). If Pcd.fc.fCompressed is 
 * one, set fc to fc / 2, and set fcPcd to fcPcd/2.
 * 4. Read a PlcBtePapx at offset FibRgFcLcb97.fcPlcfBtePapx in the Table Stream, and of size 
 * FibRgFcLcb97.lcbPlcfBtePapx. Let fcLast be the last element of plcbtePapx.aFc. If fcLast is 
 * less than or equal to fc, examine fcPcd. If fcLast is less than fcPcd, go to step 8. 
 * Otherwise, set fc to fcLast. If Pcd.fc.fCompressed is one, set fcLast to fcLast / 2. Set 
 * fcFirst to fcLast and go to step 7.
 * 5. Find the largest j such that plcbtePapx.aFc[j] ≤ fc. Read a PapxFkp at offset 
 * aPnBtePapx[j].pn *512 in the WordDocument Stream.
 * 6. Find the largest k such that PapxFkp.rgfc[k] ≤ fc. If the last element of PapxFkp.rgfc is 
 * less than or equal to fc, then cp is outside the range of character positions in this 
 * document, and is not valid. Let fcFirst be PapxFkp.rgfc[k].
 * 7. If fcFirst is greater than fcPcd, then let dfc be (fcFirst – fcPcd). If Pcd.fc.fCompressed 
 * is zero, then set dfc to dfc / 2. The first character of the paragraph is at character 
 * position PlcPcd.aCp[i] + dfc. Leave the algorithm.
 * 8. If PlcPcd.aCp[i] is 0, then the first character of the paragraph is at character position 
 * 0. Leave the algorithm.
 * 9. Set cp to PlcPcd.aCp[i]. Set i to i - 1. Go to step 2.
 *
 * To find the character position of the last character in the paragraph that contains a given 
 * character position cp:
 * 1. Follow the algorithm from Retrieving Text up to and including step 3 to find i. Also 
 * remember the FibRgFcLcb97, and PlcPcd found in step 1 of Retrieving Text. If the algorithm 
 * from Retrieving Text specifies that cp is invalid, leave the algorithm.
 * 2. Let pcd be PlcPcd.aPcd[i].
 * 3. Let fcPcd be Pcd.fc.fc. Let fc be fcPcd + 2(cp – PlcPcd.aCp[i]). Let fcMac be fcPcd + 
 * 2(PlcPcd.aCp[i+1] - PlcPcd.aCp[i]). If Pcd.fc.fCompressed is one, set fc to fc/2, set fcPcd 
 * to fcPcd /2 and set fcMac to fcMac/2.
 * 4. Read a PlcBtePapx at offset FibRgFcLcb97.fcPlcfBtePapx in the Table Stream, and of size 
 * FibRgFcLcb97.lcbPlcfBtePapx. Then find the largest j such that plcbtePapx.aFc[j] ≤ fc. If the 
 * last element of plcbtePapx.aFc is less than or equal to fc, then go to step 7. Read a PapxFkp 
 * at offset aPnBtePapx[j].pn *512 in the WordDocument Stream.
 * 5. Find largest k such that PapxFkp.rgfc[k] ≤ fc. If the last element of PapxFkp.rgfc is less 
 * than or equal to fc, then cp is outside the range of character positions in this document, 
 * and is not valid. Let fcLim be PapxFkp.rgfc[k+1].
 * 6. If fcLim ≤ fcMac, then let dfc be (fcLim – fcPcd). If Pcd.fc.fCompressed is zero, then set 
 * dfc to dfc / 2. The last character of the paragraph is at character position PlcPcd.aCp[i] + 
 * dfc – 1. Leave the algorithm.
 * 7. Set cp to PlcPcd.aCp[i+1]. Set i to i + 1. Go to step 2.
*/

int cfb_doc_parse(
		struct cfb *cfb,
		void *user_data,
		int (*text)(
			void *user_data,
			char *str
			)
		)
{
#ifdef DEBUG
	LOG("start cfb_doc_parse\n");
#endif
	
	int ret = 0;

	//Read the FIB from offset zero in the WordDocument Stream
	cfb_doc_t doc;
	ret = cfb_doc_init(&doc, cfb);
	if (ret)
		return ret;

	//get text
	_get_text_for_cp(&doc, &(doc.clx.Pcdt->PlcPcd), 0, doc.fib.rgLw97->ccpText, user_data, text);

#ifdef DEBUG
	LOG("cfb_doc_parse done\n");
#endif
	
	return 0;
}

/*
 * Main Document
 * The main document contains all content outside any of the specialized document parts, including
 * anchors that specify where content from the other document parts appears.
 * The main document begins at CP zero, and is FibRgLw97.ccpText characters long.
 * The last character in the main document MUST be a paragraph mark (Unicode 0x000D).
 */
void cfb_doc_main_document(struct cfb *cfb){
	
}


#ifdef __cplusplus
}
#endif

#endif //DOC_H_

// vim:ft=c	
