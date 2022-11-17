/**
 * File              : doc.h
 * Author            : Igor V. Sementsov <ig.kuzm@gmail.com>
 * Date              : 04.11.2022
 * Last Modified Date: 17.11.2022
 * Last Modified By  : Igor V. Sementsov <ig.kuzm@gmail.com>
 */

#ifndef DOC_H_
#define DOC_H_

#include <stdbool.h>
#ifdef __cplusplus
extern "C"{
#endif

#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
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
	uint16_t  unused0; //(2 bytes)
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
						//save (FibBase.fComplex).
	uint16_t reserved6; //(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
						//save (FibBase.fComplex).
	uint16_t reserved7; //(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
						//save (FibBase.fComplex).
	uint16_t reserved8; //(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
						//save (FibBase.fComplex).
	uint16_t reserved9; //(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
						//save (FibBase.fComplex).
	uint16_t reserved10;//(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
						//save (FibBase.fComplex).
	uint16_t reserved11;//(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
						//save (FibBase.fComplex).
	uint16_t reserved12;//(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
						//save (FibBase.fComplex).
	uint16_t reserved13;//(2 bytes): This value MUST be 0 and MUST be ignored. 
						//Word 97 and Word 2000 can put a value here when performing an incremental 
						//save (FibBase.fComplex).
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
typedef enum FibRgFcLcb {
	fibRgFcLcb97, 
	fibRgFcLcb2000, 
	fibRgFcLcb2002, 
	fibRgFcLcb2003, 
	fibRgFcLcb2007, 
} FibRgFcLcb;

struct nFib2fibRgFcLcb {
	uint16_t nFib;
	FibRgFcLcb fibRgFcLcb;
};

const struct nFib2cbRgFcLcb nFib2fibRgFcLcbTable[] = {
	{0x00C1, fibRgFcLcb97}, 
	{0x00D9, fibRgFcLcb2000}, 
	{0x0101, fibRgFcLcb2002}, 
	{0x010C, fibRgFcLcb2003}, 
	{0x0112, fibRgFcLcb2007}, 
};

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


} FibRgFcLcb2003;

#ifdef __cplusplus
}
#endif

#endif //DOC_H_
