This is an indesign script for a common production issue that occurs when the users book colours have bad contrast values and so are difficult to read when the Primary Character colour is on top of the Secondary Underline Colour

The script looks for Paragraph styles that have the underline option turned on and then checks the contrast with the character colour if it fails to meet the contrast minimum then it will adjust the tint of the underline in increments of 1% until the contrast minimum is met or the tint reaches 7.5% (current minimum but need to do more testing to find the sweet spot)
When checking the contrast between underline and character the script takes into account how the current tint is affecting the cmyk... e.g If the underline colour was 100,29,0,40 at 50% tint its cmyk values would be divided by 2 (and rounded up if .5) so we would get 50,15,0,20

The contrast check is taken from the equations from this site https://legibility.info/contrast-calculator
as far as I understand it takes the CMYK values and converts them to RGB and then find the Luminance of the two values and then calculates the contrast

* tint% is a little different from opacity as it is how much white has been put into the colour(0% tint being completely white and 100% tint being no change to the colour). Opcaity is how much of the background is allowed to show through the colour - important as not all templates have completely white backgrounds
