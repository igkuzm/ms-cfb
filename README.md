# Windows Compound Binary File Format C library

### Overview

A Compound File is made up of a number of virtual streams. These are collections of data that behave as a linear stream, although their on-disk format may be fragmented. Virtual streams can be user data, or they can be control structures used to maintain the file. Note that the file itself can also be considered a virtual stream.
All allocations of space within a Compound File are done in units called sectors. The size of a sector is definable at creation time of a Compound File, but for the purposes of this document will be 512 bytes. A virtual stream is made up of a sequence of sectors.
The Compound File uses several different types of sector: Fat, Directory, Minifat, DIF, and Storage. A separate type of 'sector' is a Header, the primary difference being that a Header is always 512 bytes long (regardless of the sector size of the rest of the file) and is always located at offset zero (0). With the exception of the header, sectors of any type can be placed anywhere within the file. The function of the various sector types is discussed below.
In the discussion below, the term SECT is used to describe the location of a sector within a virtual stream (in most cases this virtual stream is the file itself). Internally, a SECT is represented as a ULONG.
