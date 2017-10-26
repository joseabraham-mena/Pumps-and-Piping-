# Pumps and Pipes Sizing Calculator

## Background and Purpose

  I spent the summer of 2017 interning at a local company and during that time Excel VBA was a very important part. 
 That same summer I also worked on a school project that required the design of piping systems, sizing of pumps 
 and design of heat exchanger. I decided to combine both to develop a spreadsheet that could do all the background calculations
 so that the designers could just change certain variables quick and easy. 
 
  The project developed into two different spreadsheets: a technical one and a more user-friendly one. The technical spreadsheet
 is intended as a visual aid to see how diameter and material impacts pressure drop across a pipe segment and total cost. The user-
 friendly speadhseet is intended to make the design process easier to execute and understand. 

  You can download both spreadsheets here and practice their user with the problems below:
  
 * [User-friendly spreadsheet](https://joseabraham-mena.github.io/Pumps-and-Piping-/Docs/Pumps and piping calculator.xlsm ) 
 * [Technical spreadsheet](https://joseabraham-mena.github.io/Pumps-and-Piping-/Docs/Feed Section Scenario 1.xlsm) 
 

## Solved Problems

### Solved Problem 1:Finding ideal diameter given desired pressure drop and lenght

You have a pipe segment 20 meters long that runs through 2 90-degree, standard radius, fittings. 
The initial pressure at the beginning of the pipe segment is 150 kPa. 
You want the pipe segment to drop 50 kPa as it runs through the pipe. 

What is the ideal diameter?

A) Set up the spreadsheet

  1. Go to the Set Up tab
  2. Enter the number of pipe segments
  3. Click the Update button
  4. Enter the number of types of segments for each segment. In this case have two fittings
     of the same type so we enter 1.

B) Set the desired pressure drop

  5. In the next tab specify the initial and objective pressures; scroll right to
     find the cells. 
  6. Go to the initial cells and enter the specific information for the fluid in the pipe,
     an initial diameter in inches (2 for example), and enter the segment length. 

C) Solve for the ideal diameter

  7. Solve for the fanning friction factor by clicking the button. 
  8. Solve for the desired diameter.
  9. Repeat steps 7 and 8 one more time. 

### Solved Problem 2: Finding pressure drop given diameter and length

You have been given a pipe loop with two segments. The first runs for 5 meters and through a 2 inch
pipe. The second segment funs for 2 meters through a half inch pipe. 

Find the pressure drop at the end of the pipe. 




## Built With

  * EXCEL VBA 2015

## Authors

  * Jose Abraham Mena 

## Acknowledgments

  * Benjamin Feikima
  * Zac Ardavanis
  * Professor Steven Pohler
