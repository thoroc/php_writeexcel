<?php
/*
 * Copyleft 2002 Johann Hanne
 *
 * This is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This software is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this software; if not, write to the
 * Free Software Foundation, Inc., 59 Temple Place,
 * Suite 330, Boston, MA  02111-1307 USA
 */

namespace writeexcel\classes;

/**
 * This is the Spreadsheet::WriteExcel Perl package ported to PHP
 * Spreadsheet::WriteExcel was written by John McNamara, jmcnamara@cpan.org
 */
class OLEWriter
{
    var $OLEfilename;
    var $OLEtmpfilename; /* ABR */
    var $fileHandle;
    var $fileClosed;
    var $internalFileHandle;
    var $BIFFOnly;
    var $sizeAllowed;
    var $BIFFSize;
    var $bookSize;
    var $bigBlocks;
    var $listBlocks;
    var $rootStart;
    var $blockCount;

    /**
     * Constructor
     */
    public function __construct( $filename )
    {
        $this->OLEfilename = $filename;
        $this->fileHandle = false;
        $this->fileClosed = 0;
        $this->internalFileHandle = 0;
        $this->BIFFOnly = 0;
        $this->sizeAllowed = 0;
        $this->BIFFSize = 0;
        $this->bookSize = 0;
        $this->bigBlocks = 0;
        $this->listBlocks = 0;
        $this->rootStart = 0;
        $this->blockCount = 4;

        $this->initialize();
    }
    /*
     * Check for a valid filename and store the filehandle.
     */

    function initialize()
    {
        $OLEfile = $this->OLEfilename;

        /* Check for a filename. Workbook.pm will catch this first. */
        if( $OLEfile == '' )
        {
            trigger_error( "Filename required", E_USER_ERROR );
        }

        /*
         * If the filename is a resource it is assumed that it is a valid
         * filehandle, if not we create a filehandle.
         */
        if( is_resource( $OLEfile ) )
        {
            $fh = $OLEfile;
        }
        else
        {
            // Create a new file, open for writing
            $fh = fopen( $OLEfile, "wb" );
            // The workbook class also checks this but something may have
            // happened since then.
            if( !$fh )
            {
                trigger_error( "Can't open $OLEfile. It may be in use or " .
                        "protected", E_USER_ERROR );
            }

            $this->internalFileHandle = 1;
        }

        // Store filehandle
        $this->fileHandle = $fh;
    }

    /**
     * Set the size of the data to be written to the OLE stream
     *
     * $big_blocks = (109 depot block x (128 -1 marker word)
     *               - (1 x end words)) = 13842
     * $maxsize    = $big_blocks * 512 bytes = 7087104
     */
    public function setSize( $size )
    {
        $maxsize = 7087104;

        if( $size > $maxsize )
        {
            trigger_error( "Maximum file size, $maxsize, exceeded. To create " .
                    "files bigger than this limit please use the " .
                    "workbookbig class.", E_USER_ERROR );
            return ($this->sizeAllowed = 0);
        }

        $this->BIFFSize = $size;

        // Set the min file size to 4k to avoid having to use small blocks
        if( $size > 4096 )
        {
            $this->bookSize = $size;
        }
        else
        {
            $this->bookSize = 4096;
        }

        return ($this->sizeAllowed = 1);
    }

    /**
     * Calculate various sizes needed for the OLE stream
     */
    public function calculateSizes()
    {
        $datasize = $this->bookSize;

        if( $datasize % 512 == 0 )
        {
            $this->bigBlocks = $datasize / 512;
        }
        else
        {
            $this->bigBlocks = floor( $datasize / 512 ) + 1;
        }
        // There are 127 list blocks and 1 marker blocks for each big block
        // depot + 1 end of chain block
        $this->listBlocks = floor( ($this->bigBlocks) / 127 ) + 1;
        $this->rootStart = $this->bigBlocks;

        //print $this->_biffsize.    "\n";
        //print $this->_big_blocks.  "\n";
        //print $this->_list_blocks. "\n";
    }

    /**
     * Write root entry, big block list and close the filehandle.
     * This method must be called so that the file contents are
     * actually written.
     */
    public function close()
    {

        if( !$this->sizeAllowed )
        {
            return;
        }

        if( !$this->BIFFOnly )
        {
            $this->writePadding();
            $this->writePropertyStorage();
            $this->writeBigBlockDepot();
        }

        // Close the filehandle if it was created internally.
        if( $this->internalFileHandle )
        {
            fclose( $this->fileHandle );
        }
        /* ABR */
        if( $this->OLEtmpfilename != '' )
        {
            $fh = fopen( $this->OLEtmpfilename, "rb" );
            if( $fh == false )
            {
                trigger_error( "Can't read temporary file.", E_USER_ERROR );
            }
            fpassthru( $fh );
            fclose( $fh );
            unlink( $this->OLEtmpfilename );
        };

        $this->fileClosed = 1;
    }

    /**
     * Write BIFF data to OLE file.
     */
    public function write( $data )
    {
        fputs( $this->fileHandle, $data );
    }

    /**
     * Write OLE header block.
     */
    public function writeHeader()
    {
        if( $this->BIFFOnly )
        {
            return;
        }

        $this->calculateSizes();

        $root_start = $this->rootStart;
        $num_lists = $this->listBlocks;

        $id = pack( "C8", 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 );
        $unknown1 = pack( "VVVV", 0x00, 0x00, 0x00, 0x00 );
        $unknown2 = pack( "vv", 0x3E, 0x03 );
        $unknown3 = pack( "v", -2 );
        $unknown4 = pack( "v", 0x09 );
        $unknown5 = pack( "VVV", 0x06, 0x00, 0x00 );
        $num_bbd_blocks = pack( "V", $num_lists );
        $root_startblock = pack( "V", $root_start );
        $unknown6 = pack( "VV", 0x00, 0x1000 );
        $sbd_startblock = pack( "V", -2 );
        $unknown7 = pack( "VVV", 0x00, -2, 0x00 );
        $unused = pack( "V", -1 );

        fputs( $this->fileHandle, $id );
        fputs( $this->fileHandle, $unknown1 );
        fputs( $this->fileHandle, $unknown2 );
        fputs( $this->fileHandle, $unknown3 );
        fputs( $this->fileHandle, $unknown4 );
        fputs( $this->fileHandle, $unknown5 );
        fputs( $this->fileHandle, $num_bbd_blocks );
        fputs( $this->fileHandle, $root_startblock );
        fputs( $this->fileHandle, $unknown6 );
        fputs( $this->fileHandle, $sbd_startblock );
        fputs( $this->fileHandle, $unknown7 );

        for( $c = 1; $c <= $num_lists; $c++ )
        {
            $root_start++;
            fputs( $this->fileHandle, pack( "V", $root_start ) );
        }

        for( $c = $num_lists; $c <= 108; $c++ )
        {
            fputs( $this->fileHandle, $unused );
        }
    }

    /**
     * Write big block depot.
     */
    function writeBigBlockDepot()
    {
        $numBlocks = $this->bigBlocks;
        $numLists = $this->listBlocks;
        $totalBlocks = $numLists * 128;
        $usedBlocks = $numBlocks + $numLists + 2;

        $marker = pack( "V", -3 );
        $endOfChain = pack( "V", -2 );
        $unused = pack( "V", -1 );

        for( $i = 1; $i <= ($numBlocks - 1); $i++ )
        {
            fputs( $this->fileHandle, pack( "V", $i ) );
        }

        fputs( $this->fileHandle, $endOfChain );
        fputs( $this->fileHandle, $endOfChain );

        for( $c = 1; $c <= $numLists; $c++ )
        {
            fputs( $this->fileHandle, $marker );
        }

        for( $c = $usedBlocks; $c <= $totalBlocks; $c++ )
        {
            fputs( $this->fileHandle, $unused );
        }
    }

    /**
     * Write property storage.
     * @todo add summary sheets
     */
    public function writePropertyStorage()
    {
        $rootsize = -2;
        $booksize = $this->bookSize;

        //                name          type  dir start  size
        $this->writePPS( 'Root Entry', 0x05, 1, -2, 0x00 );
        $this->writePPS( 'Book', 0x02, -1, 0x00, $booksize );
        $this->writePPS( '', 0x00, -1, 0x00, 0x0000 );
        $this->writePPS( '', 0x00, -1, 0x00, 0x0000 );
    }

    /**
     * Write property sheet in property storage
     */
    private function writePPS( $name, $type, $dir, $start, $size )
    {
        $names = array();
        $length = 0;

        if( $name != '' )
        {
            $name = $name . "\0";
            // Simulate a Unicode string
            $chars = preg_split( "''", $name, -1, PREG_SPLIT_NO_EMPTY );
            foreach( $chars as $char )
            {
                array_push( $names, ord( $char ) );
            }
            $length = strlen( $name ) * 2;
        }

        $rawname = call_user_func_array( 'pack', array_merge( array( "v*" ), $names ) );
        $zero = pack( "C", 0 );

        $pps_sizeofname = pack( "v", $length );   //0x40
        $pps_type = pack( "v", $type );     //0x42
        $pps_prev = pack( "V", -1 );        //0x44
        $pps_next = pack( "V", -1 );        //0x48
        $pps_dir = pack( "V", $dir );      //0x4c

        $unknown1 = pack( "V", 0 );

        $pps_ts1s = pack( "V", 0 );         //0x64
        $pps_ts1d = pack( "V", 0 );         //0x68
        $pps_ts2s = pack( "V", 0 );         //0x6c
        $pps_ts2d = pack( "V", 0 );         //0x70
        $pps_sb = pack( "V", $start );    //0x74
        $pps_size = pack( "V", $size );     //0x78

        fputs( $this->fileHandle, $rawname );
        fputs( $this->fileHandle, str_repeat( $zero, (64 - $length ) ) );
        fputs( $this->fileHandle, $pps_sizeofname );
        fputs( $this->fileHandle, $pps_type );
        fputs( $this->fileHandle, $pps_prev );
        fputs( $this->fileHandle, $pps_next );
        fputs( $this->fileHandle, $pps_dir );
        fputs( $this->fileHandle, str_repeat( $unknown1, 5 ) );
        fputs( $this->fileHandle, $pps_ts1s );
        fputs( $this->fileHandle, $pps_ts1d );
        fputs( $this->fileHandle, $pps_ts2d );
        fputs( $this->fileHandle, $pps_ts2d );
        fputs( $this->fileHandle, $pps_sb );
        fputs( $this->fileHandle, $pps_size );
        fputs( $this->fileHandle, $unknown1 );
    }

    /**
     * Pad the end of the file
     */
    private function writePadding()
    {
        $BIFFsize = $this->BIFFSize;

        if( $BIFFsize < 4096 )
        {
            $minSize = 4096;
        }
        else
        {
            $minSize = 512;
        }

        if( $BIFFsize % $minSize != 0 )
        {
            $padding = $minSize - ($BIFFsize % $minSize);
            fputs( $this->fileHandle, str_repeat( "\0", $padding ) );
        }
    }
}