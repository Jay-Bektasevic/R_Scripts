

#     =================================================================================================    
#     Script Name:  R_2_Excel_Tabs.R 
#     =================================================================================================    
#
#     Create Date:  2018-08-06
#     Author:       Jay Bektasevic
#     Author Email: jbektasevic@clearpathmutual.com
# 
#     Description:  This script's primary purpose is to take a dataset (import or direct fetch from the 
#                   database), then split it up by an accident year and export it to an excel workbook 
#                   where each tab represents an Accident Year. It saves a ton of manual labor copying 
#                   and pasting. 
#                   Although the request came from CNA (an excess insurer for the period of 1994 - 2006), 
#                   with some easy tweaks, this script can be used for other purposes, especially, if it 
#                   involves a large number of records which need to be split up into different tabs of  
#                   the same workbook.  
# 
# 
#     Call by:      [Application Name]
#                   [Job]
#                   [Interface]
# 
# 
#     Used By:      Mainly Accounting, but anyone can use it. 
#                 
#     =================================================================================================
#                                       SUMMARY OF CHANGES
#     -------------------------------------------------------------------------------------------------
#     Date:                   Author:           Comments:
#     ----------------------- ----------------- -------------------------------------------------------
#       
#     =================================================================================================


#   Increase the heap size before the rJava package is loaded. rJava is a dependency for "xlsx" package.
#   Heap Memory means in programming, an area of memory reserved for data that is created at runtime that is, 
#   when the program actually executes. In contrast, the stack is an area of memory used for data whose 
#   size can be determined when the program is compiled.
#   Java heap is the heap size allocated to JVM applications which takes care of the new objects being created. 
#   If the objects being created exceed the heap size, it will throw an error saying memory Out of Bound.

    options(java.parameters = "-Xmx16000m") # Thats 16GB
    
    library(xlsx)
    library(progress)
    library(readxl)

    setwd("K:/Departments/Accounting/Jay Bektasevic")

#   Either import a file or execute a pass-through SQL query via RODBC package.
    CNA_Loss <- read_excel("CNA_loss_request.xlsx", sheet = "Sheet1")
    
#   Select columns from the dataset 
    input_vars <-c( "ACCIDENT_YEAR","CLAIM_NUMBER","CLAIMANT","ADJUSTER","INSURED_NAME","INJURY_DATE","CLAIM_STATUS_CODE",
                    "MEDICAL_PAID_BEFORE_REC","INDEMNITY_PAID_BEFORE_REC","EXPENSE_PAID_BEFORE_REC","LEGAL_PAID_BEFORE_REC",
                    "PAID_TOTAL_BEFORE_REC","MEDICAL_RECOVERY","INDEMNITY_RECOVERY","EXPENSE_RECOVERY","LEGAL_RECOVERY",
                    "RECOVERY_TOTAL","MEDICAL_REINSURANCE_RECOVERY","INDEMNITY_REINSURANCE_RECOVERY",
                    "EXPENSE_REINSURANCE_RECOVERY","LEGAL_REINSURANCE_RECOVERY","REINSURANCE_RECOVERY_TOTAL",
                    "MEDICAL_RESERVE","INDEMNITY_RESERVE","EXPENSE_RESERVE","LEGAL_RESERVE","RESERVE_TOTAL",
                    "MEDICAL_INCURRED","INDEMNITY_INCURRED","EXPENSE_INCURRED","LEGAL_INCURRED","INCURRED_TOTAL"
    )
#   Check to see if we get the negated columns
    CNA_Loss[, !names(CNA_Loss) %in% input_vars] 
    
#   Fix the NULL's for aesthetic purposes. Since, NULL's here are char format by replacing it with NA it will allow us to remove them during export. 
    CNA_Loss[CNA_Loss$ADJUSTER == "NULL", "ADJUSTER" ] <- NA 
    
#   Check the counts of claims by Status Code.    
    summary(as.factor(CNA_Loss$CLAIM_STATUS_CODE))
    
#   Fix the status code to list the full description
    CNA_Loss$CLAIM_STATUS_CODE[CNA_Loss$CLAIM_STATUS_CODE == "CLOS"] <- "CLOSED"
    CNA_Loss$CLAIM_STATUS_CODE[CNA_Loss$CLAIM_STATUS_CODE == "RECL"] <- "RECLOSED"
    CNA_Loss$CLAIM_STATUS_CODE[CNA_Loss$CLAIM_STATUS_CODE == "REOP"] <- "REOPEN"

#   Fix the date format - Timestamp format causes issues while exporting to Excel.
    CNA_Loss$INJURY_DATE <-
        as.Date(CNA_Loss$INJURY_DATE, format = "%Y%m%d")

#   Break up the file and export to tabs by Accident Year.

    write_to_excel <- function(ACCIDENT_YEAR, dataset) {
       
#       Add 8 empty rows to the top for aesthetic purposes.
#   ---------------------------------------------------------------------------------------------------
        pad <- dataset[FALSE, ]
        pad[nrow(pad) + 8, ] <- NA
        pad2 <- dataset[dataset$ACCIDENT_YEAR == ACCIDENT_YEAR,]
        data <- rbind(pad, pad2)
#   ---------------------------------------------------------------------------------------------------
        write.xlsx(
            as.data.frame(data),
            file = "CNA_Loss_Run_Final.xlsx",
            sheetName = paste0("AY_", ACCIDENT_YEAR),
            row.names = FALSE,
            append = TRUE,
            showNA = FALSE # Removes NA's 
        )
        
    }

#   A function that will demand a garbage collection in each iteration of the loop.
    
    jgc <- function() {
        .jcall("java/lang/System", method = "gc")
    }

#   Run the Losses and export to Excel.
#   ---------------------------------------------------------------------------------------------------
#   Progress bar enviroment
    pb <- progress_bar$new(
        format = "Exporting [:bar] :percent in :elapsed",
        total = length(unique(CNA_Loss$ACCIDENT_YEAR)) - 1,
        clear = FALSE,
        width = 60
    )

    for (i in min(CNA_Loss$ACCIDENT_YEAR):(max(CNA_Loss$ACCIDENT_YEAR) - 1)) {
        
        jgc()
        write_to_excel(ACCIDENT_YEAR = i, dataset = CNA_Loss[, input_vars])
        pb$tick()
        if (i == (max(CNA_Loss$ACCIDENT_YEAR) - 1))
            cat(paste0(" Export Complete!\n Location of the exported file is: ", getwd()))
    }

#   ---------------------------------------------------------------------------------------------------