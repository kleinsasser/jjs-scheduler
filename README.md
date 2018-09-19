# Scheduler
Program that takes employee information (15 variables) (ex. available hours, maximum hours, worker type) from an excel document and returns a one-week schedule based on the provided information. UI was entirely neglected, made for an ex-boss and my own personal gratification.

Instructions:

System Requirements
-System must have Microsoft Excel with a version from no later than 2007 installed.
-System must have the free software “R” installed, can be downloaded from this link {https://cran.r-project.org/mirrors.html} (select link with location closest to your area, select download for your operating system)

LEAVE THE SCHEDULER FOLDER IN ITS CURRENT STATE SAVED TO THE DESKTOP

Step 1 - Fill in “SchedulerData.xlsx” with appropriate information.

DO NOT EDIT COLUMNS OR ROWS SHADED IN GRAY OR CHANGE COLUMN NAMES
If an employee is terminated or added, just insert and delete rows accordingly. (Do not leave rows blank)

Variable Information - Employees Sheet
This is where you fill in information about your employees’ availability and desired hours.

Name - Name of employee 
Driver - If the employee is a driver or can drive, insert TRUE, otherwise FALSE
InShop - Same thing as Driver but for InShop employees
Days - TRUE or FALSE on whether or not the employee works days
Nights - TRUE or FALSE on whether or not the employee works nights
MaxHours - The maximum amount of hours the employee can work
TargetHoursDay - The desired amount of day shift hours desired for an employee
TargetHoursNight - The desired amount of night shift hours desired for an employee
TargetHoursInShop - The desired amount of in-shop hours desired for an employee
TargetHoursDriver - The desired amount of driving hours desired for an employee
PriorityDriver - Gives priority of driver shift assignment to the employee with the lowest value. Each employee must have a number, 1, 2, 3, or 4, assigned.
PriorityInShop - Gives priority of driver shift assignment to the employee with the lowest value. Each employee must have a number, 1, 2, 3, or 4, assigned.

WedAM->TuesPM - This is where specific shifts requested off are accounted for—insert TRUE if a shift has been requested off. Color code cells to ease the process if desired. Copy and paste color-coded cells to the corresponding sections of the schedule template to avoid color-coding more than once.

Variable Information - Days/Nights Sheets
This is where you will create shifts that need to be assigned.

Shift - Descriptive but short name of the shift specifying its duration. Each shift name must be unique. Example set of shifts: 10-5, 11-R, 11-R(2)-indicating a shift that has one or more duplicates, 10-5(D)-indicating a driving shift
Hours - Number indicating the length in hours of the particular shift.
Driver - TRUE or FALSE indicating whether or not the shift is a driving shift (does not eliminate the necessity of a unique shift name)

SAVE INFORMATION IN FILE BEFORE MOVING ON

Step 2 - Generate Schedule in R

If this is not the first schedule you’ve generated, move the previous “schedule.xlsx” and “employeeHours.xlsx” files out of the “Scheduler” folder.

Open up the R Application, copy and paste the following chunk of code into the console (after the “>”), and press enter/return:

setwd('~/Desktop/Scheduler')
source('~/Desktop/Scheduler/generateSchedule.R')
generateSchedule()

Two files: “schedule.xlsx” and “employeeHours.xlsx” should have appeared in the Scheduler folder.
“schedule.xlsx” is the actual schedule that will be pasted into the format template
“employeeHours.xlsx” is a file describing the amount of hours each employee was scheduled by “Day”, “Night”, “InShop”, and “Driver”. Should be used to make sure all employees are scheduled properly. If the schedule is for just about any reason inadequate, it can most likely be solved by adjusting TargetHours of employees, saving the file, running the R code again, and reopening the schedule and employeeHours files.

Step 3 - Format Schedule

Using the “formatTemplate.xlsx” file, manually schedule manager and PIC shifts, copy and paste color coded shifts into the employee area, and copy and paste the schedule from “schedule.xlsx” into the format template.

Simple as that.
