resetSchedule <- function(employeeStats) {
  Employees <<- read.xlsx("SchedulerData.xlsx", sheetIndex = 1, stringsAsFactors = FALSE)
  AM <<- read.xlsx("SchedulerData.xlsx", sheetIndex = 2, stringsAsFactors = FALSE)
  PM <<- read.xlsx("SchedulerData.xlsx", sheetIndex = 3, stringsAsFactors = FALSE)
}


assignShift <- function(AM_PM, start_end, shift, driver) {
  
  # only take those employees who haven't worked
  E <<- E[which(E$Worked == FALSE),]
  
  # subset for AM or PM shift
  if (AM_PM == "AM") {
    availableE <<- subset(E, Days == TRUE)
  } else if (AM_PM == "PM") {
    availableE <<- subset(E, Nights == TRUE)
  } else {
    stop("'AM' or 'PM' for AM_PM argument")
  }
  
  # subset for drivers
  if (driver == TRUE) {
    availableE <<- subset(availableE, Driver == TRUE)
  } else {
    availableE <<- subset(availableE, InShop == TRUE)
  }
  
  # assign on shift sheet and adjust hours for employee
  if (AM_PM == "AM") {
    A <<- subset(availableE, availableE$MaxHours > (availableE$TotalHours + as.integer(AM[which(AM$Shift == start_end), "Hours"])))
    
    if (driver == TRUE) {
      A$DiffType <<- A$TotalHoursDriver-A$TargetHoursDriver
      A$DiffTime <<- A$TotalHoursDay-A$TargetHoursDay
      A <<- A[order(A$PriorityDriver, A$DiffTime, A$DiffType, -A$TargetHoursDriver, -A$TargetHoursDay),]
    } else {
      A$DiffType <<- A$TotalHoursInShop-A$TargetHoursInShop
      A$DiffTime <<- A$TotalHoursDay-A$TargetHoursDay
      A <<- A[order(A$PriorityInShop, A$DiffTime, A$DiffType, -A$TargetHoursInShop, -A$TargetHoursDay),]
    }
    
    A2 <<- A[which(A$MaxHours > A$TotalHours + 5),]
    
    worker <<- A2[1,]
    
    AM[which(AM$Shift == start_end), shift] <<- worker$Name
    Employees[which(Employees$Name == worker$Name), "TotalHours"] <<- Employees[which(Employees$Name == worker$Name), "TotalHours"] + as.integer(AM[which(AM$Shift == start_end), "Hours"])
    Employees[which(Employees$Name == worker$Name), "TotalHoursDay"] <<- Employees[which(Employees$Name == worker$Name), "TotalHoursDay"] + as.integer(AM[which(AM$Shift == start_end), "Hours"])
    E[which(E$Name == worker$Name), "Worked"] <<- TRUE
    
    if (driver == TRUE) {
      Employees[which(Employees$Name == worker$Name), "TotalHoursDriver"] <<- Employees[which(Employees$Name == worker$Name), "TotalHoursDriver"] + as.integer(AM[which(AM$Shift == start_end), "Hours"])
    } else {
      Employees[which(Employees$Name == worker$Name), "TotalHoursInShop"] <<- Employees[which(Employees$Name == worker$Name), "TotalHoursInShop"] + as.integer(AM[which(AM$Shift == start_end), "Hours"])
    }
    
    E[which(E$Name == "Not Scheduled1"), "Worked"] <<- FALSE
  }
  
  if (AM_PM == "PM") {
    A <<- subset(availableE, availableE$MaxHours > (availableE$TotalHours + as.integer(PM[which(PM$Shift == start_end), "Hours"])))
    
    if (driver == TRUE) {
      A$DiffType <<- A$TotalHoursDriver-A$TargetHoursDriver
      A$DiffTime <<- A$TotalHoursNight-A$TargetHoursNight
      A <<- A[order(A$PriorityDriver, A$DiffTime, A$DiffType, -A$TargetHoursDriver, -A$TargetHoursNight),]
    } else {
      A$DiffType <<- A$TotalHoursInShop-A$TargetHoursInShop
      A$DiffTime <<- A$TotalHoursDay-A$TargetHoursDay
      A <<- A[order(A$PriorityInShop, A$DiffTime, A$DiffType, -A$TargetHoursInShop, -A$TargetHoursNight),]
    }
    
    A2 <<- A[which(A$MaxHours > A$TotalHours + 5),]
    
    worker <<- A2[1,]
    
    PM[which(PM$Shift == start_end), shift] <<- worker$Name
    Employees[which(Employees$Name == worker$Name), "TotalHours"] <<- Employees[which(Employees$Name == worker$Name), "TotalHours"] + as.integer(PM[which(PM$Shift == start_end), "Hours"])
    Employees[which(Employees$Name == worker$Name), "TotalHoursNight"] <<- Employees[which(Employees$Name == worker$Name), "TotalHoursNight"] + as.integer(PM[which(PM$Shift == start_end), "Hours"])
    E[which(E$Name == worker$Name), "Worked"] <<- TRUE
    
    if (driver == TRUE) {
      Employees[which(Employees$Name == worker$Name), "TotalHoursDriver"] <<- Employees[which(Employees$Name == worker$Name), "TotalHoursDriver"] + as.integer(PM[which(PM$Shift == start_end), "Hours"])
    } else {
      Employees[which(Employees$Name == worker$Name), "TotalHoursInShop"] <<- Employees[which(Employees$Name == worker$Name), "TotalHoursInShop"] + as.integer(PM[which(PM$Shift == start_end), "Hours"])
    }
    
    E[which(E$Name == "Not Scheduled"), "Worked"] <<- FALSE
  }
}


scheduleShift <<- function(shiftsheet, shift, AM_PM) {
  
  Off <<- paste0(shift, AM_PM)
  E <<- Employees[which(Employees[Off] == FALSE),]
  
  for (i in shiftsheet$Shift) {
    assignShift(AM_PM = AM_PM, start_end = i, shift = shift, driver = shiftsheet[which(shiftsheet$Shift == i),"Driver"])
  }
}

scheduleWeek <<- function(shiftsheet, AM_PM) {
  for (i in subset(colnames(shiftsheet), colnames(shiftsheet) != "Shift" & colnames(shiftsheet) != "Hours" & colnames(shiftsheet) != "Driver")) {
    scheduleShift(shiftsheet = shiftsheet, shift = i, AM_PM = AM_PM)
  }
  
}

reformat <<- function() {
  schedule <<- data.frame(as.character(Employees$Name), NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA)
  colnames(schedule) <<- c("Name","WedAM", "WedPM", "ThuAM", "ThuPM", "FriAM", "FriPM", "SatAM", "SatPM", "SunAM", "SunPM", "MonAM", "MonPM", "TuesAM", "TuesPM")
  
  for (i in colnames(AM)[4:10]) {
    for (j in AM$Shift) {
      name <- AM[which(AM$Shift == j), i]
      day <- paste0(i, "AM")
      shift <- j
      
      schedule[which(schedule$Name == name), day] <<- shift
    }
  }
  
  for (i in colnames(PM)[4:10]) {
    for (j in PM$Shift) {
      name <- PM[which(PM$Shift == j), i]
      day <- paste0(i, "PM")
      shift <- j
      
      schedule[which(schedule$Name == name), day] <<- shift
    }
  }
}

generateSchedule <- function() {
  xlsxinstalled <- any(row.names(installed.packages()) == "xlsx")
  
  if (xlsxinstalled == FALSE) {
    install.packages("xlsx")
  }
  
  library(xlsx)
  
  resetSchedule()
  scheduleWeek(AM, "AM")
  scheduleWeek(PM, "PM")
  reformat()
  
  HoursTable <<- data.frame(Employees$Name, Employees$TotalHoursDay, Employees$TotalHoursNight, Employees$TotalHoursInShop, Employees$TotalHoursDriver, Employees$TotalHours, Employees$Name)
  colnames(HoursTable) <<- c("Name", "Day", "Night", "InShop", "Driver", "Total", "Name")
  
  write.xlsx(schedule, "schedule.xlsx", sheetName = "Schedule", showNA = FALSE, row.names = FALSE)
  write.xlsx(HoursTable, "employeeHours.xlsx", sheetName = "EmployeeHours", row.names = FALSE)
}
