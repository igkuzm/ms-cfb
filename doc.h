/**
 * File              : doc.h
 * Author            : Igor V. Sementsov <ig.kuzm@gmail.com>
 * Date              : 04.11.2022
 * Last Modified Date: 15.11.2022
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
 * function `property_set_get`
 * Read properties from standard common property set file, execute callback for each found
 * property. Return error code. To stop function execution you may return non 0 in callback
 * function. 
 */
static int 
property_set_get(FILE * fp, void * user_data,
	int (*callback)(void * user_data, uint32_t propid, uint32_t dwType, uint8_t * value));


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
	                       //Stream. An STSH that specifies the style sheet for this document begins at 
						   //this offset. 
	uint32_t lcbStshf;     //(4 bytes): An unsigned integer that specifies the size, in bytes, of the 
						   //STSH that begins at offset fcStshf in the Table Stream. This MUST be a 
						   //nonzero value. 
	uint32_t fcPlcffndRef; //(4 bytes): An unsigned integer that specifies an offset in the Table 
						   //Stream. A PlcffndRef begins at this offset and specifies the locations of 
						   //footnote references in the Main Document, and whether those references use 
						   //auto-numbering or custom symbols. If lcbPlcffndRef is zero, fcPlcffndRef 
						   //is undefined and MUST be ignored. 
	uint32_t lcbPlcffndRef;//(4 bytes): An unsigned integer that specifies the size, in bytes, of the 
						   //PlcffndRef that begins at offset fcPlcffndRef in the Table Stream. 
} FibRgFcLcb97;


#ifdef __cplusplus
}
#endif

#endif //DOC_H_
