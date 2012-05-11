SxBRC Ordnance Survey Grid Converter
====================================

## Description

SxBRC OS Grid Converter is a simple Excel add-in that adds functions for conversion of Ordnance Survey grid references (e.g., TQ2113) to easting and northing Cartesian coordinates.

## Excel Functions

Grid Converter adds four functions to Excel:

* EASTING
* NORTHING
* EASTING_C
* NORTHING_C

The two functions suffixed with `_C` will move the generated coordinates to the centre of the referenced OS grid square; the other two functions place the coordinates at the bottom-left of the grid square.

## Syntax

    EASTING(text)

Same syntax applies to the other functions.

## Examples

If cell A1 contains the reference TQ1234:

    Formula         Result
    =EASTING(A1)    512000
    =NORTHING(A1)   134000
    =EASTING_C(A1)  512500
    =NORTHING_C(A1) 134500

The conversion alogrithm can also handle the [DINTY](http://en.wikipedia.org/wiki/DINTY) tetrad reference style:

    Formula         Result
    =EASTING(A1)    526000
    =NORTHING(A1)   112000

## Contact

This add-in was created by Charles Roper (charlesroper) for the [Sussex Biodiversity Record Centre](http://sxbrc.org.uk).

