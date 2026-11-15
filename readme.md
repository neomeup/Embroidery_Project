Embroidery Template is a fun project to turn photos into embroidery templates.

Run using make sheet after placing the photo that you want to build a template from inside of the input photo directory.

The script will output a xlxs file to output templates.

Use master color list to input the colors that you own if you want the script to build the template only from those colors.

You can also modify the quantity of colors used if you would like to run using all possible DMC colors or if you own many colors and would just like to limit the amount you use. 

Some notes are that if you run with all possible colors (thus with USE_ALL_COLORS set to TRUE), even for relatively simple designs, the script could take a minute or two to run.
If you run with a smaller quantity of colors on a normal size grid, the script should take less than 15-20 seconds. 

Also, there is a known issue where red parts of input photos sometimes are output blue.  This is certainly an issue with the color calculation, you could use optimization or play around with it to improve that.