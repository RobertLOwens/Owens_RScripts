fuzzyLogic <- function(data1, data2, col.x = "default", col.y = "default", 
                       fuzziness = 0.92, error = .02, othercolumns = "", 
                       ex_file = "default", append = FALSE) {
    
# Author: Robert Owens 
# Created: 1/15/2019
# Purpose: Join two tables together which don't have a primary key to join on. 

# This function will take two data frames that you want to join.
# It will do a fuzzy join on the two tables using the col.x and col.y args
# and a given probability "percentage". The output of this function will be a 
# venndiagram, a list of items around the given percentage, a list of duplicates
# and a final report. When using "jw", use "percentage" as an argument, when 
# using general, use "generaldist". 

# generalmatch <- stringdist_left_join(x, y, by = c(name = "name2"),
#                      distance_col = "distance", max_dist = generaldist,
#                      ignore_case = TRUE)

library(fuzzyjoin)
library(dplyr)
library(readxl)
library(openxlsx)
library(VennDiagram)


# Validation
if (!is.numeric(fuzziness) | fuzziness > 1 & fuzziness < 0) {
  stopifnot("Invalid fuzziness: Needs to a numeric value between 0 and 1")
} else if (!is.numeric(error) | error > 1 & error < 0) {
  stopifnot("Invalid error value: Needs to be a numeric value between 0 and 1")
} else if ((!is.character(col.x) | !is.character(col.y)) &
           !is.na(colnames(data1)[col.x]) | !is.na(colnames(data1)[col.y])) {
    stopifnot(paste("Invalid col.x or col.y: either not a character, or does",
                    "not exist in the dataframe"))
} else if (!is.data.frame(data1) | !is.data.frame(data2)) {
  stopifnot("Invalid Data Format: Data needs to be in data table format")
} else if (append == TRUE & !exists(x = "ex_file"))
  stopifnot("Append is TRUE, however the file arg was not defined")

colnames(data1)[colnames(data1) == col.x] <- "ID"
colnames(data2)[colnames(data2) == col.y] <- "ID2"

# Fuzzy Join
fuzzylogic_df <- stringdist_left_join(data1, data2, by = c(ID = "ID2"), 
                                      method = "jw", p = .1, 
                                      distance_col = "distance", 
                                      ignore_case = TRUE)

fuzzylogic_df$distance <- round(1 - fuzzylogic_df$distance, digits = 2)

if (exists("othercolumns")) {
  filter_fuzzylogic_df <- select(fuzzylogic_df, ID, ID2, distance, othercolumns)
} else {
  filter_fuzzylogic_df <-select(fuzzylogic_df, ID, ID2, distance)
}

filter_fuzzylogic_df <- filter_fuzzylogic_df %>%
  filter(distance >= (fuzziness)) %>%
  arrange(distance)

# Venn Diagram Summary

intersect <- summarise(filter_fuzzylogic_df, n = n())
intersect$Data <- "Matched"
intersect <- intersect[c(2,1)]
colnames(intersect) <- c("Data", "Observations")

grid.newpage()
venndiagram <- draw.pairwise.venn(area1 = nrow(data1),
                                  area2 = nrow(data2),
                                  cross.area = intersect$Observations,
                                  category = c("Data #1", "Data #2"),
                                  fill = c("light blue", "pink"),
                                  alpha = 0.6,
                                  lty = rep("blank", 2),
                                  cat.pos = c(0, 0),
                                  cat.dist = rep(0.025, 2),
                                  scaled = FALSE,
                                  euler.d = FALSE
)

# List all duplicates
duplicates <- filter_fuzzylogic_df %>%
  group_by(ID2) %>%
  summarise(n = n()) %>%
  filter(n > 1)

colnames(duplicates) <- c("Name", "Number of Duplications")

# %Error 
percent_error <- fuzzylogic_df %>%
  filter(distance <= fuzziness + error & distance >= fuzziness - error) %>%
  mutate(IsIncluded = ifelse(distance >= fuzziness, "Included", 
                             "Not Included")) %>%
  select(ID, ID2, distance, IsIncluded) %>%
  arrange(distance)

# Create a new tab on the file specified
if (append == TRUE) {
  wb <- loadWorkbook(ex_file)
  addWorksheet(wb, "Final Report")
  writeData(wb, sheet = "Final Report", x = filter_fuzzylogic_df)
  saveWorkbook(wb, ex_file, overwrite = TRUE)
}


results <- list(filter_fuzzylogic_df, duplicates, percent_error, venndiagram)
names(results) <- c("Final Report", "Duplicates", "Percent_Error",
                    "Venn Diagram")
return(results)
}

fuzzyLogicExcel <- function(ex_file = "default", col.x = "default", 
                            col.y = "default", fuzziness = 0.92, error = .02, 
                            othercolumns = "", append = FALSE) {
    
# Author: Robert Owens
# Created: 1/15/2019
# Purpose: Join two tables together from an Excel Sheet
# which don't have a primary key to join on. 

# This function will take one excel sheet with two tabs that you want to join.
    
library(readxl)

#Import the data
data1 <- read_xlsx(ex_file, col_names = TRUE, sheet = 1)
data2 <- read_xlsx(ex_file, col_names = TRUE, sheet = 2)

results <- fuzzyLogic(data1 = data1, data2 = data2, col.x = col.x,
                      col.y = col.y, fuzziness = fuzziness, error = error,
                      othercolumns = othercolumns, append = append)

return(results)
}

# Test Data
# data1 <- data.frame(lastname = c("Owens", "Jewell", "Eng", "Brown"),
#                     firstname = c("Robert","Gregory", "Timothy", "Mark"),
#                     ID = c("Owens, Robert", "Jewell, Gregory", "Eng, Timothy",
#                            "Brown, Mark"))
# data2 <- data.frame(lastname = c("Owens", "Jewell", "Eng"), #, "Snippy"),
#                    firstname = c("Rob","Greg", "Tim"), #, "Washigton"),
#                    ID = c("Owens, Rob", "Jewell, Greg", "Eng, Tim"),
#                           #"Washington, Snippy"),
#                    number = c(1, 2, 3))
