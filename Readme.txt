As mentioned in the PSC Readme file, the difficult part with the GradientFill API is its unsigned integers. 

The TRIVERTEX documentation  states The color information of each channel is specified as a value from 0x0000 to 0xff00.", or in VB terminology "from &H0000 to &HFF00". This means from 0 to 65280. We all know that assignig a value greater than 32767 to a VB integer results in run-time error 6 : overflow.

I assume you are familiar with the GDI API's and structures. This is NOT a tutorial. If you are want to know and understand, please consult your documentation. If desperation creeps in, feel free to email me at paul_turcksin@hotmail.com.

So, the trick is to bypass VB internal checking. The easiest way is to assign a hex value
   trivertex(0).Red = &HAF00

but this is hardcoded i.e. not very flexible. It is however possible (and has been done) to manipulate strings to obtain this flexibity.  I find the code complicated and hard to understand (Yes, I'm lazy and YES this is my personal view).

You can also use some mathematics to turn the "greater than 32767" value into a negative number. The API expects bits, not a value. I tried this approach but it is cumbersome. (You have to do you calculations in a Long). Anyway, the code - and the results - didn't satisfy me.

The solution I finally came up with is to move bytes in memory. This also bypasses VB checking and offers simplicity and clarity.

   Dim arBytes(7) as bytes   ->   bytes per integer, 4 integers (red,green,blue and apha)

   arBytes(0) = 0 to 255   ->the red component
   arBytes(3)= idem for the green component
   arBytes(5)   ... blue
   arBytes(7)   ..alpha

then you copy these to the Trivertex structure in one go
   (CopyMemory destination,source,nmbr of bytes        (CopyMery is an alias for RtlMoveMemory)
    CopyMemory trivertex(0).red, arBytes(0),8

and there you go   ... gradient fill.

Greetings to all of you.