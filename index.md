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
   
 ![Image of Steps 1-4](https://joseabraham-mena.github.io/Pumps-and-Piping-/Docs/Example problems/Problem 1/2,3,4PNG.PNG)  

B) Set the desired pressure drop

  5. In the Pressure Drops in Loops tab specify the fluid properties.
  6. Choose the pipe material from the drop down menu, enter a guess diameter value
     and the pipe length. 
     
 ![Image of Steps 5-6](https://joseabraham-mena.github.io/Pumps-and-Piping-/Docs/Example problems/Problem 1/5,6.PNG)
  
  7. Scroll rigth to the Segment Summary Table and enter the initial pressure.
  8. In the Solve for Specific Pressure Drop Table enter the objective pressure.
  
![Image of Steps 7-8](https://joseabraham-mena.github.io/Pumps-and-Piping-/Docs/Example problems/Problem 1/7,8.PNG)

C) Solve for the ideal diameter

  9.  Solve for the fanning friction factor by clicking the button. 
  10. Solve for the desired diameter.
  11. Repeat steps 7 and 8 one more time. 

### Solved Problem 2: Finding pressure drop given diameter and length

You have been given a pipe loop with two segments. The first runs for 5 meters and through a 2 inch
pipe. The second segment funs for 2 meters through a half inch pipe. 

Find the pressure drop at the end of the pipe. 

A) Repeat steps 1-5 in from Problem 2.

B) In the Pressure Drops in Loops tab specify the pipe material, diameter and length

C) Find the pressure drop

   1. Click the button to solve for the Fanning Friction Factor
![Image of step 1](https://joseabraham-mena.github.io/Pumps-and-Piping-/Docs/Example problems/Problem 2/1.PNG)   

   Because the two segments are continous, we say that the pressure into the second segment is equal to the
   pressure outside of the first segment

   2. Scroll right to the Segment Summary table and read the value of Pressure 2(kPa)
   
 ![Image of step 2](https://joseabraham-mena.github.io/Pumps-and-Piping-/Docs/Example problems/Problem 2/2.PNG)

## Built With

  * EXCEL VBA 2015

## Authors

  * Jose Abraham Mena 

## Acknowledgments

  * Benjamin Feikima
  * Zac Ardavanis
  * Professor Steven Pohler
