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

/*
 * This is the Spreadsheet::WriteExcel Perl package ported to PHP
 * Spreadsheet::WriteExcel was written by John McNamara, jmcnamara@cpan.org
 */

class Format
{
    var $xfIndex;
    var $fontIndex;
    var $font;
    var $size;
    var $bold;
    var $italic;
    var $color;
    var $underline;
    var $fontStrikeout;
    var $fontOutline;
    var $fontShadow;
    var $fontScript;
    var $fontFamily;
    var $fontCharset;
    var $numFormat;
    var $hidden;
    var $locked;
    var $textHalign;
    var $textWrap;
    var $textValign;
    var $textJustLast;
    var $rotation;
    var $FgColor;
    var $BgColor;
    var $pattern;
    var $bottom;
    var $top;
    var $left;
    var $right;
    var $bottomColor;
    var $topColor;
    var $leftColor;
    var $rightColor;

    /**
     * Constructor
     */
    public function __construct()
    {
        $_ = func_get_args();

        $this->xfIndex = (sizeof( $_ ) > 0) ? array_shift( $_ ) : 0;

        $this->fontIndex = 0;
        $this->font = 'Arial';
        $this->size = 10;
        $this->bold = 0x0190;
        $this->italic = 0;
        $this->color = 0x7FFF;
        $this->underline = 0;
        $this->fontStrikeout = 0;
        $this->fontOutline = 0;
        $this->fontShadow = 0;
        $this->fontScript = 0;
        $this->fontFamily = 0;
        $this->fontCharset = 0;

        $this->numFormat = 0;

        $this->hidden = 0;
        $this->locked = 1;

        $this->textHalign = 0;
        $this->textWrap = 0;
        $this->textValign = 2;
        $this->textJustLast = 0;
        $this->rotation = 0;

        $this->FgColor = 0x40;
        $this->BgColor = 0x41;

        $this->pattern = 0;

        $this->bottom = 0;
        $this->top = 0;
        $this->left = 0;
        $this->right = 0;

        $this->bottomColor = 0x40;
        $this->topColor = 0x40;
        $this->leftColor = 0x40;
        $this->rightColor = 0x40;

        // Set properties passed to writeexcel_workbook::addformat()
        if( sizeof( $_ ) > 0 )
        {
            call_user_func_array( array( &$this, 'set_properties' ), $_ );
        }
    }

    /**
     * Copy the attributes of another writeexcel_format object.
     */
    public function copy( $other )
    {
        $xf = $this->xfIndex;   // Backup XF index
        foreach( $other as $key->$value )
        {
            $this->{$key} = $value;
        }
        $this->xfIndex = $xf;   // Restore XF index
    }

    /**
     * Generate an Excel BIFF XF record.
     */
    public function getXf()
    {
        $_ = func_get_args();

        // $record    Record identifier
        // $length    Number of bytes to follow
        // $ifnt      Index to FONT record
        // $ifmt      Index to FORMAT record
        // $style     Style and other options
        // $align     Alignment
        // $icv       fg and bg pattern colors
        // $fill      Fill and border line style
        // $border1   Border line style and color
        // $border2   Border color
        // Set the type of the XF record and some of the attributes.
        if( $_[0] == "style" )
        {
            $style = 0xFFF5;
        }
        else
        {
            $style = $this->locked;
            $style |= $this->hidden << 1;
        }

        // Flags to indicate if attributes have been set.
        $atr_num = ($this->numFormat != 0) ? 1 : 0;
        $atr_fnt = ($this->fontIndex != 0) ? 1 : 0;
        $atr_alc = $this->textWrap ? 1 : 0;
        $atr_bdr = ($this->bottom ||
                $this->top ||
                $this->left ||
                $this->right) ? 1 : 0;
        $atr_pat = ($this->FgColor != 0x41 ||
                $this->BgColor != 0x41 ||
                $this->pattern != 0x00) ? 1 : 0;
        $atr_prot = 0;

        // Reset the default colors for the non-font properties
        if( $this->FgColor == 0x7FFF ) $this->FgColor = 0x40;
        if( $this->BgColor == 0x7FFF ) $this->BgColor = 0x41;
        if( $this->bottomColor == 0x7FFF ) $this->bottomColor = 0x41;
        if( $this->topColor == 0x7FFF ) $this->topColor = 0x41;
        if( $this->leftColor == 0x7FFF ) $this->leftColor = 0x41;
        if( $this->rightColor == 0x7FFF ) $this->rightColor = 0x41;

        // Zero the default border colour if the border has not been set.
        if( $this->bottom == 0 )
        {
            $this->bottomColor = 0;
        }
        if( $this->top == 0 )
        {
            $this->topColor = 0;
        }
        if( $this->right == 0 )
        {
            $this->rightColor = 0;
        }
        if( $this->left == 0 )
        {
            $this->leftColor = 0;
        }

        // The following 2 logical statements take care of special cases in
        // relation to cell colors and patterns:
        // 1. For a solid fill (_pattern == 1) Excel reverses the role of
        //    foreground and background colors
        // 2. If the user specifies a foreground or background color
        //    without a pattern they probably wanted a solid fill, so we
        //    fill in the defaults.
        if( $this->pattern <= 0x01 &&
                $this->BgColor != 0x41 &&
                $this->FgColor == 0x40 )
        {
            $this->FgColor = $this->BgColor;
            $this->BgColor = 0x40;
            $this->pattern = 1;
        }

        if( $this->pattern <= 0x01 &&
                $this->BgColor == 0x41 &&
                $this->FgColor != 0x40 )
        {
            $this->BgColor = 0x40;
            $this->pattern = 1;
        }

        $record = 0x00E0;
        $length = 0x0010;

        $ifnt = $this->fontIndex;
        $ifmt = $this->numFormat;

        $align = $this->textHalign;
        $align |= $this->textWrap << 3;
        $align |= $this->textValign << 4;
        $align |= $this->textJustLast << 7;
        $align |= $this->rotation << 8;
        $align |= $atr_num << 10;
        $align |= $atr_fnt << 11;
        $align |= $atr_alc << 12;
        $align |= $atr_bdr << 13;
        $align |= $atr_pat << 14;
        $align |= $atr_prot << 15;

        $icv = $this->FgColor;
        $icv |= $this->BgColor << 7;

        $fill = $this->pattern;
        $fill |= $this->bottom << 6;
        $fill |= $this->bottomColor << 9;

        $border1 = $this->top;
        $border1 |= $this->left << 3;
        $border1 |= $this->right << 6;
        $border1 |= $this->topColor << 9;

        $border2 = $this->leftColor;
        $border2 |= $this->rightColor << 7;

        $header = pack( "vv", $record, $length );
        $data = pack( "vvvvvvvv", $ifnt, $ifmt, $style, $align, $icv, $fill, $border1, $border2 );

        return($header . $data);
    }

    /**
     * Generate an Excel BIFF FONT record.
     */
    public function getFont()
    {
        // $record     Record identifier
        // $length     Record length
        // $dyHeight   Height of font (1/20 of a point)
        // $grbit      Font attributes
        // $icv        Index to color palette
        // $bls        Bold style
        // $sss        Superscript/subscript
        // $uls        Underline
        // $bFamily    Font family
        // $bCharSet   Character set
        // $reserved   Reserved
        // $cch        Length of font name
        // $rgch       Font name

        $dyHeight = $this->size * 20;
        $icv = $this->color;
        $bls = $this->bold;
        $sss = $this->fontScript;
        $uls = $this->underline;
        $bFamily = $this->fontFamily;
        $bCharSet = $this->fontCharset;
        $rgch = $this->font;

        $cch = strlen( $rgch );
        $record = 0x31;
        $length = 0x0F + $cch;
        $reserved = 0x00;

        $grbit = 0x00;

        if( $this->italic )
        {
            $grbit |= 0x02;
        }

        if( $this->fontStrikeout )
        {
            $grbit |= 0x08;
        }

        if( $this->fontOutline )
        {
            $grbit |= 0x10;
        }

        if( $this->fontShadow )
        {
            $grbit |= 0x20;
        }

        $header = pack( "vv", $record, $length );
        $data = pack( "vvvvvCCCCC", $dyHeight, $grbit, $icv, $bls, $sss, $uls, $bFamily, $bCharSet, $reserved, $cch );

        return($header . $data . $this->font);
    }

    /**
     * Returns a unique hash key for a font.
     * Used by writeexcel_workbook::_store_all_fonts()
     */
    public function getFontKey()
    {
        # The following elements are arranged to increase the probability of
        # generating a unique key. Elements that hold a large range of numbers
        # eg. _color are placed between two binary elements such as _italic
        #
        $key = $this->font . $this->size .
                $this->fontScript . $this->underline .
                $this->fontStrikeout . $this->bold . $this->fontOutline .
                $this->fontFamily . $this->fontCharset .
                $this->fontShadow . $this->color . $this->italic;

        $key = preg_replace( '/ /', '_', $key ); # Convert the key to a single word

        return $key;
    }

    /**
     * Returns the used by Worksheet->_XF()
     */
    public function getXfIndex()
    {
        return $this->xfIndex;
    }

    /**
     * Used in conjunction with the set_xxx_color methods to convert a color
     * string into a number. Color range is 0..63 but we will restrict it
     * to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
     */
    public function getColor( $color = false )
    {
        $colors = array(
            'aqua' => 0x0F,
            'cyan' => 0x0F,
            'black' => 0x08,
            'blue' => 0x0C,
            'brown' => 0x10,
            'magenta' => 0x0E,
            'fuchsia' => 0x0E,
            'gray' => 0x17,
            'grey' => 0x17,
            'green' => 0x11,
            'lime' => 0x0B,
            'navy' => 0x12,
            'orange' => 0x35,
            'purple' => 0x14,
            'red' => 0x0A,
            'silver' => 0x16,
            'white' => 0x09,
            'yellow' => 0x0D
        );

        // Return the default color, 0x7FFF, if undef,
        if( $color === false )
        {
            return 0x7FFF;
        }

        // or the color string converted to an integer,
        if( isset( $colors[strtolower( $color )] ) )
        {
            return $colors[strtolower( $color )];
        }

        // or the default color if string is unrecognised,
        if( preg_match( '/\D/', $color ) )
        {
            return 0x7FFF;
        }

        // or an index < 8 mapped into the correct range,
        if( $color < 8 )
        {
            return $color + 8;
        }

        // or the default color if arg is outside range,
        if( $color > 63 )
        {
            return 0x7FFF;
        }

        // or an integer in the valid range
        return $color;
    }

    /**
     * Set cell alignment.
     */
    public function setAlign( $location )
    {
        // Ignore numbers
        if( preg_match( '/\d/', $location ) )
        {
            return;
        }

        $location = strtolower( $location );

        switch( $location )
        {
            case 'left':
                $this->setTextHalign( 1 );
                break;

            case 'centre':
            case 'center':
                $this->setTextHalign( 2 );
                break;

            case 'right':
                $this->setTextHalign( 3 );
                break;

            case 'fill':
                $this->setTextHalign( 4 );
                break;

            case 'justify':
                $this->setTextHalign( 5 );
                break;

            case 'merge':
                $this->setTextHalign( 6 );
                break;

            case 'equal_space':
                $this->setTextHalign( 7 );
                break;

            case 'top':
                $this->setTextValign( 0 );
                break;

            case 'vcentre':
            case 'vcenter':
                $this->setTextValign( 1 );
                break;
                break;

            case 'bottom':
                $this->setTextValign( 2 );
                break;

            case 'vjustify':
                $this->setTextValign( 3 );
                break;

            case 'vequal_space':
                $this->setTextValign( 4 );
                break;
        }
    }

    /**
     * Set vertical cell alignment. This is required by the set_properties()
     * method to differentiate between the vertical and horizontal properties.
     */
    public function setValign( $location )
    {
        $this->setAlign( $location );
    }

    /**
     * This is an alias for the unintuitive set_align('merge')
     */
    public function setMerge()
    {
        $this->setTextHalign( 6 );
    }

    /**
     * Bold has a range 0x64..0x3E8.
     * 0x190 is normal. 0x2BC is bold.
     */
    public function setBold( $weight = 1 )
    {
        if( $weight == 1 )
        {
            // Bold text
            $weight = 0x2BC;
        }

        if( $weight == 0 )
        {
            // Normal text
            $weight = 0x190;
        }

        if( $weight < 0x064 )
        {
            // Lower bound
            $weight = 0x190;
        }

        if( $weight > 0x3E8 )
        {
            // Upper bound
            $weight = 0x190;
        }

        $this->bold = $weight;
    }

    /**
     * Set all cell borders (bottom, top, left, right) to the same style
     */
    public function setBorder( $style )
    {
        $this->setBottom( $style );
        $this->setTop( $style );
        $this->setLeft( $style );
        $this->setRight( $style );
    }

    /**
     * Set all cell borders (bottom, top, left, right) to the same color
     */
    public function setBorderColor( $color )
    {
        $this->setBottomColor( $color );
        $this->setTopColor( $color );
        $this->setLeftColor( $color );
        $this->setRightColor( $color );
    }

    /**
     * Convert hashes of properties to method calls.
     */
    public function setProperties()
    {
        $_ = func_get_args();

        $properties = array();
        foreach( $_ as $props )
        {
            if( is_array( $props ) )
            {
                $properties = array_merge( $properties, $props );
            }
            else
            {
                $properties[] = $props;
            }
        }

        foreach( $properties as $key => $value )
        {

            // Strip leading "-" from Tk style properties eg. -color => 'red'.
            $key = preg_replace( '/^-/', '', $key );

            /* Make sure method names are alphanumeric characters only, in
              case tainted data is passed to the eval(). */
            if( preg_match( '/\W/', $key ) )
            {
                trigger_error( "Unknown property: $key.", E_USER_ERROR );
            }

            /* Evaling all $values as a strings gets around the problem of
              some numerical format strings being evaluated as numbers, for
              example "00000" for a zip code. */
            if( is_int( $key ) )
            {
                eval( "\$this->set_$value();" );
            }
            else
            {
                eval( "\$this->set_$key('$value');" );
            }
        }
    }

    public function setFont( $font )
    {
        $this->font = $font;
    }

    public function setSize( $size )
    {
        $this->size = $size;
    }

    public function setItalic( $italic = 1 )
    {
        $this->italic = $italic;
    }

    public function setColor( $color )
    {
        $this->color = $this->getColor( $color );
    }

    public function setUnderline( $underline = 1 )
    {
        $this->underline = $underline;
    }

    public function setFontStrikeout( $fontStrikeout = 1 )
    {
        $this->fontStrikeout = $fontStrikeout;
    }

    public function setFontOutline( $fontOutline = 1 )
    {
        $this->fontOutline = $fontOutline;
    }

    public function setFontShadow( $font_shadow = 1 )
    {
        $this->fontShadow = $font_shadow;
    }

    public function setFontScript( $font_script = 1 )
    {
        $this->fontScript = $font_script;
    }
    /* Undocumented */

    public function setFontFamily( $font_family = 1 )
    {
        $this->fontFamily = $font_family;
    }
    /* Undocumented */

    public function setFontCharset( $font_charset = 1 )
    {
        $this->fontCharset = $font_charset;
    }

    public function setNumFormat( $num_format = 1 )
    {
        $this->numFormat = $num_format;
    }

    public function setHidden( $hidden = 1 )
    {
        $this->hidden = $hidden;
    }

    public function setLocked( $locked = 1 )
    {
        $this->locked = $locked;
    }

    public function setTextHalign( $align )
    {
        $this->textHalign = $align;
    }

    public function setTextWrap( $wrap = 1 )
    {
        $this->textWrap = $wrap;
    }

    public function setTextValign( $align )
    {
        $this->textValign = $align;
    }

    public function setTextJustLast( $textJustLast = 1 )
    {
        $this->textJustLast = $textJustLast;
    }

    public function setRotation( $rotation = 1 )
    {
        $this->rotation = $rotation;
    }

    public function setFgColor( $color )
    {
        $this->FgColor = $this->getColor( $color );
    }

    public function setBgColor( $color )
    {
        $this->BgColor = $this->getColor( $color );
    }

    public function setPattern( $pattern = 1 )
    {
        $this->pattern = $pattern;
    }

    public function setBottom( $bottom = 1 )
    {
        $this->bottom = $bottom;
    }

    public public function setTop( $top = 1 )
    {
        $this->top = $top;
    }

    public function setLeft( $left = 1 )
    {
        $this->left = $left;
    }

    public function setRight( $right = 1 )
    {
        $this->right = $right;
    }

    public function setBottomColor( $color )
    {
        $this->bottomColor = $this->getColor( $color );
    }

    public function setTopColor( $color )
    {
        $this->topColor = $this->getColor( $color );
    }

    public function setLeftColor( $color )
    {
        $this->leftColor = $this->getColor( $color );
    }

    public function setRightColor( $color )
    {
        $this->rightColor = $this->getColor( $color );
    }
}