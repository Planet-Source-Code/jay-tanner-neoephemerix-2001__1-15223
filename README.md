﻿<div align="center">

## NeoEphemerix\_2001

<img src="PIC20012111932345726.gif">
</div>

### Description

<BR><BR>NeoEphemerix 2001 - v1 Beta 2<BR><BR>

Written using VB 6<BR>

Requires 800x600 display or better<BR><BR>

<BR><BR>

This program is for astronomy hobbyists who want to create their own custom VB astronomical almanac program. It is a very complex program consisting hundreds of thousands of orbital computations and represents about 6 month's work so far.<BR><BR>

The program will generate high-precision ephemerides for the sun and planets from Mercury to Neptune.<BR><BR>

It has reached the beta level of functionality and I encourage any fellow astro-computationists to give it a try and offer any comments, bug reports and suggestions regarding it.<BR><BR>

Anyone who ever wanted to learn how to perform their own high-precision planetary orbit computations, may find the source code helpful, but the math required is rather advanced.<BR><BR>

A new version with even more computations is in the works. Any suggestions from users of this program will be considered in the design of future upgrades.<BR><BR>

NOTE:

Due to the enormous size of the mathematical core modules and the complexity of the computations required to get almanic-like accuracy, the program takes about 30 minutes to compile into an executable on a 200 MHz machine and will produce a finished program about 4.3 megs in size.

The visual magnitude computation of the planet Saturn will be applied to a future version, but magnitudes are computed for the other planets. The allowance for the rings contribution to its brightness hasn't been formulated yet.<BR><BR>

This version of the program does not yet specifically check to see if the date entered is in the proper range for the selected planet, so the following table is provided as a guide.<BR><BR>

<BR>

VSOP87 Heliocentric coordinates are theoretically accurate to an arcsecond or better within the following ranges:<BR>

Mercury to Mars - 2000 BC to 6000 AD<BR>

Jupiter and Saturn - 1 BC to 4000 AD<BR>

Uranus and Neptune - 6000 BC to 8000 AD<BR><BR>

To any users familiar with the astronomical algorithms of Jean Meeus and others, this program applies many of the same concepts, but at a higher level of precision not possible from the limited tables applied in the popular books on astronomical computing.<BR><BR>

It is based on a Visual BASIC implementation of the full VSOP87 theory of planetary orbits in spherical variables.<BR><BR>

Its theoretical heliocentric accuracy is to within ±1 arcsecond or better over the ranges specified for each planet in terms of dynamical time.<BR><BR>

Since the full theory is implemented, the accuracy of the orbit computations compares very favorably with the published almanacs.<BR><BR>

To achieve this level of accuracy, over 30,000 computational terms are applied to the raw, dynamical orbit computations.<BR><BR>

The computations include corrections for precession and the long-term effects of relativity on the orbits. Then corrections are applied for light-time, aberration, reduction to the standard FK5 system of coordinates and nutation.<BR><BR><BR>

FEATURES INCLUDE:<BR><BR>

Both VSOP87 heliocentric and apparent geocentric ecliptical and equatorial coordinates<BR><BR>

Ephemerides tables can be generated by the day, hour or minute<BR> and can be saved to disk as plain text files<BR><BR>

Allowance for delta-T can be applied when known.<BR>

Distances to the planets from the sun or Earth can be displayed in astronomical units, millions of kilometers or miles.<BR><BR>

Hour angles may be displayed in hours minutes and seconds, decimal hours, degrees minutes and seconds of arc or decimal degrees.<BR><BR>

Latitudes may be displayed in degrees minutes and seconds of arc or decimal degrees.<BR><BR>

A table showing the VSOP heliocentric position and geocentric ecliptical and equatorial coordinates for the sun and the eight major planets at any given moment can be displayed<BR><BR>

Basic astronomical data such as the mean and apparent obliquity of the ecliptic, mean and apparent sidereal time at Greenwich and nutation in longitude can also be computed.<BR><BR>

The program also has a stay-on-top feature that can be toggled to lock the window in front of other windows when needed.<BR><BR>

All program settings are preserved when the program terminates and are recalled the next time the program starts up.<BR>

<BR><BR>

This program is still a rough draft but functional enough to be useful.<BR><BR>

I would appreciate any feedback on user of this code who are also into astro-computing, since it would help me to improve on future implementations.

<BR><BR>
 
### More Info
 
Date, time, mode selections

Creates a high-precision almanac of planet positions.


<span>             |<span>
---                |---
**Submitted On**   |2001-01-31 02:58:00
**By**             |[Jay Tanner](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jay-tanner.md)
**Level**          |Advanced
**User Rating**    |5.0 (130 globes from 26 users)
**Compatibility**  |VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD148352112001\.zip](https://github.com/Planet-Source-Code/jay-tanner-neoephemerix-2001__1-15223/archive/master.zip)







