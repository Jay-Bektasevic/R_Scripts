
#     =================================================================================================    
#     Script Name:  Read_all_sheets.R
#     =================================================================================================    
#
#     Create Date:  2019-05-07
#     Author:       Jay Bektasevic
#     Author Email: jbektasevic@clearpathmutual.com
# 
#     Description:  Reads in multiple sheets form the same workbook and merges them together in a 
#                   single data frame. 
#                   
#                   This is nice and compact, and most of all its lightning fast compared to other methods.
# 
# 
#     Call by:      [Application Name]
#                   [Job]
#                   [Interface]
# 
# 
#     Used By:      Anyone
#                 
#     =================================================================================================
#                                       SUMMARY OF CHANGES
#     -------------------------------------------------------------------------------------------------
#     Date:                   Author:           Comments:
#     ----------------------- ----------------- -------------------------------------------------------
#     2019-07-03              Jay Bektasevic    Make sure to update file directory and file name. 
#     =================================================================================================

#   List of packages to load:
    packages <- c('readxl',
                  'dplyr',
                  'data.table'
    )
    
    new_packages <-
        packages[!(packages %in% installed.packages()[, "Package"])]
    
    if (length(new_packages))
        install.packages(new_packages)
    
#   Load Neccessary Packages
    sapply(packages, require, character.only = TRUE)   


    setwd("K:/Departments/Accounting/Actuarial Data/2019 Q3 Loss Analysis")
    
   
    input_files <- list.files(  pattern="*.xlsx") 
    
    inputWB <- input_files[2]
    
#   This is set to read all sheets, but, you can specify which ones to read in by means of specifying 
#   the index of a particular sheet or a sequence of sheets.
    sheetsToRead <- excel_sheets(inputWB) 
    
    
#   sheetsToRead <- excel_sheets(inputWB)[1:length(excel_sheets(inputWB))-1] 
    listOfSheets <- lapply(sheetsToRead, function(x,inputWB) read_excel(inputWB, sheet=x), inputWB) 
    
#   Bind the list and convert to data.table/data.frame
    final_dataFrame <- rbindlist(listOfSheets, idcol= NULL)

#   Export to .csv for further manipulation/visulation. Use fwrite() function very fast compare to write.csv() function.
    data.table::fwrite(final_dataFrame, file = "K:/Departments/Accounting/Actuarial Data/2019 Q3 Loss Analysis/Rating_Detail_EARNED_09_30_2019.csv", row.names = FALSE)
    

    
    
#   Quick test for aggreagate losses by AY
    x <- final_dataFrame %>%
        select(
            AY,
            INCURRED_TOTAL
        ) %>%
        group_by(AY) %>%
        summarise(
            total_incurred = sum(INCURRED_TOTAL)
        )
    
#   ____________________________________________________________________________________________________
#   This is where I used custom function to read in sheet by sheet and get the totals.
#   This can be utilized for other purposes as well. 
    
    inputWB <- "Rating Details Exhibit_EARNED_9_5_18.xlsx"
    
#   This is set to read all sheets, but, you can specify which ones to read in by means of specifying 
#   the index of a particular sheet or a sequence of sheets.
    sheetsToRead <- excel_sheets(inputWB) 
    
    
    read_ep_sheets <- function(file_to_read = inputWB,
                               sheet = 1) {
        rd_sheet <- read_excel(file_to_read, sheet = sheetsToRead[sheet])
        rd_sheet$CY <-
            substr(sheetsToRead[sheet], 4, 7)
        
        
        rd <- rd_sheet %>%
            select(
                CY,
                POLICY_YEAR,
                EXPOSURE_AMOUNT,
                MANUAL_PREMIUM,
                POLICY_PREMIUM,
                EARNED_PAYROLL,
                EARNED_PREMIUM
            ) %>%
            group_by(CY,
                     POLICY_YEAR) %>%
            summarise(
                EXPOSURE_AMOUNT = sum(EXPOSURE_AMOUNT),
                MANUAL_PREMIUM  = sum(MANUAL_PREMIUM),
                POLICY_PREMIUM  = sum(POLICY_PREMIUM),
                EARNED_PAYROLL  = sum(EARNED_PAYROLL),
                EARNED_PREMIUM  = sum(EARNED_PREMIUM)
            )
        
        return(rd)
        
    }
    
    
    #   read in all sheets
    
    
    # Run the get_losses() function and export to Excel.
    rd_all <- NULL
    
    t0 <- Sys.time()
    
    for (i in 1:length(sheetsToRead)) {
        print(
            paste0(
                "Please Wait...Now Fetching Calendar Year: ",
                substr(sheetsToRead[i], 4, 7),
                " Rating Detail...Run on: ",
                Sys.time()
            )
        )
        
        x <-  read_ep_sheets(file_to_read = inputWB, sheet = i)
        
        rd_all <- rbind(rd_all, x)
        
        if (i == length(sheetsToRead))
            cat(paste0(" Import Complete!\n"))
    }
    
    t1 <- Sys.time()
    t1 - t0
    
#   Compare new to old file to make sure that nothing out of ordinary is happening.  
    rd_all_new$EXPOSURE_AMOUNT_DIFF <- (rd_all_new$EXPOSURE_AMOUNT - rd_all$EXPOSURE_AMOUNT)
    rd_all_new$MANUAL_PREMIUM_DIFF <- (rd_all_new$MANUAL_PREMIUM - rd_all$MANUAL_PREMIUM)
    rd_all_new$POLICY_PREMIUM_DIFF <- (rd_all_new$POLICY_PREMIUM - rd_all$POLICY_PREMIUM)
    rd_all_new$EARNED_PAYROLL_DIFF <- (rd_all_new$EARNED_PAYROLL - rd_all$EARNED_PAYROLL)
    rd_all_new$EARNED_PREMIUM_DIFF <- (rd_all_new$EARNED_PREMIUM - rd_all$EARNED_PREMIUM)
    
    
 
    
    