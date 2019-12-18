
#     =================================================================================================    
#     Script Name:  Read_multiple_workbooks.R
#     =================================================================================================    
#
#     Create Date:  2018-12-18
#     Author:       Jay Bektasevic
#     Author Email: jbektasevic@clearpathmutual.com
# 
#     Description:  Reads in multiple workbooks and merges them together in a 
#                   single data frame. 
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
#     2018-12-18  
#     =================================================================================================

#   Increase the heap size before the rJava package is loaded - it's a dependency for "xlsx" package.
    options(java.parameters = "-Xmx16000m")

#   List of packages to load:
    packages <- c('readxl',
                  'dplyr',
                  'xlsx',
                  "tidyr"
    )
    
    new_packages <-
        packages[!(packages %in% installed.packages()[, "Package"])]

    if (length(new_packages))
        install.packages(new_packages)

#   Load Neccessary Packages
    sapply(packages, require, character.only = TRUE)

    setwd("K:/Departments/Accounting/Management Information Analyses/Claim Triangles/New Triangle Data")
    
#   List of all excel files

    file.list <- list.files(pattern='*.xlsx') 
    file.list 



#   Dates to be used for as of stamp
    dates <- c("03/31/2006", 	"06/30/2006", 	"09/30/2006", 	"12/31/2006", 	
               "03/31/2007", 	"06/30/2007", 	"09/30/2007", 	"12/31/2007", 	
               "03/31/2008", 	"06/30/2008", 	"09/30/2008", 	"12/31/2008", 	
               "03/31/2009", 	"06/30/2009", 	"09/30/2009", 	"12/31/2009", 	
               "03/31/2010", 	"06/30/2010", 	"09/30/2010", 	"12/31/2010", 	
               "03/31/2011",    "06/30/2011", 	"09/30/2011", 	"12/31/2011", 	
               "03/31/2012", 	"06/30/2012", 	"09/30/2012", 	"12/31/2012", 	
               "03/31/2013", 	"06/30/2013", 	"09/30/2013", 	"12/31/2013", 	
               "03/31/2014", 	"06/30/2014", 	"09/30/2014", 	"12/31/2014", 	
               "03/31/2015", 	"06/30/2015", 	"09/30/2015",   "12/31/2015",	
               "03/31/2016",	"06/30/2016",	"09/30/2016",	"12/31/2016",	
               "03/31/2017",	"06/30/2017",	"09/30/2017",	"12/31/2017",
               "03/31/2018",	"06/30/2018",	"09/30/2018",	"12/31/2018"

    )
    
    yr_end_dt <- c("12/31/2006",
                   "12/31/2007",
                   "12/31/2008",
                   "12/31/2009",
                   "12/31/2010",
                   "12/31/2011",
                   "12/31/2012",
                   "12/31/2013",
                   "12/31/2014",
                   "12/31/2015",
                   "12/31/2016",
                   "12/31/2017",
                   "12/31/2018"
    )

    columns <- c("INSURED_NO",	"CLAIM_NUMBER",	"CLAIM_STATUS",	"CLAIM_TYPE_GROUP",	"ACCIDENT_DATE",	"ACCIDENTYEAR",	
                 "POLICY_YEAR",	"MEDNETPAID",	"INDNETPAID",	"EXPNETPAID",	"MEDOPENRESERVE",	"INDOPENRESERVE",	
                 "EXPOPENRESERVE",	"MEDNETINCURRED",	"INDNETINCURRED",	"EXPNETINCURRED",	"TOTALNETPAID",	"TOTALOPENRESERVE",	
                 "TOTALNETINCURRED"
    )
    
#   Create a function to calculate a number of months between two dates. 
#   I chose to do this in native R code so that I don't have to load an extra library (lubridate).
    
    elapsed_months <- function(end_date, start_date) {
        ed <- as.POSIXlt(end_date)
        sd <- as.POSIXlt(start_date)
        12 * (ed$year - sd$year) + (ed$mon - sd$mon) +1
    }
    
 
    
#   Initilize dfs object to store read-in files
    dfs <- NULL

#   Loop through the file.list and read them in one by one - and bind them into dfs object.
    for (i in 1:50) {
        gc()
        print(paste0("Please Wait...Now Importing Workbook - '", file.list[i],"'"))
        df <- read_excel(file.list[i], sheet ="SQL Results")
        df$asOf <- as.Date(dates[i], "%m/%d/%Y")
        dfs <- rbind(dfs,df)
    }
#   Remove the Insured_Name and claimant name columns for formatting purposes.   
    dfs <- dfs[,-c(2,7)]

#   Export results to .csv (Very Fast!)
    data.table::fwrite(dfs, file = "claims.csv", row.names = FALSE)

#   Export to excel (If the data frame contains more then 100k records - this will not work!).    
    write.xlsx(
        dfs,
        file = "claims.xlsx",
        sheetName = "triangles",
        row.names = FALSE,
        append = TRUE
    )
   
#     =================================================================================================    
#                                               TESTING For tringles 
#     =================================================================================================   
     
   create_df <- function(file.list.index = 1, dates.index = 1){
#   Read in Excel file for selected columns.    
    df <- read_excel(file.list[file.list.index], sheet ="SQL Results")[columns]

#   Convert       
    df$CLAIM_NUMBER <- as.numeric(df$CLAIM_NUMBER)
   
#   Notification only claims that need to be bined with rest of the claims.        
    noti_only <- read_excel("BREEZE_NOTIFICATION_ONLY.xlsx")
    noti_only$CLAIM_TYPE_NEW <-  ifelse((noti_only$TOTALNETINCURRED > 0 ), "LOST_TIME", "MED_ONLY")
    
    noti_only$CLAIM_STATUS_NEW <- ifelse((noti_only$CLAIM_STATUS == "CLOS" 
                                         | noti_only$CLAIM_STATUS == "PEND" 
                                         | noti_only$CLAIM_STATUS == "RECL"), "CLOSED", "OPEN")
    noti_claims <- noti_only %>%
        mutate(ASOF = as.Date(ASOF)) %>% 
        filter(ASOF <= as.Date(dates[dates.index], "%m/%d/%Y")) %>%  # As of date to match the claims file
        select(-ASOF) # remove AsOF column
    
  
    
    
#   Segragate claims based on criteria  
    # df$claim_type <- ifelse((df$CLAIM_TYPE_GROUP == "M" 
    #                              & df$INDNETINCURRED == 0 
    #                              & df$EXPNETINCURRED == 0), "MED_ONLY",  "LOST_TIME")
    
    df$CLAIM_TYPE_NEW <- ifelse((df$INDNETINCURRED > 0 ), "LOST_TIME", "MED_ONLY")
   
    
    df$CLAIM_STATUS_NEW <- ifelse((df$CLAIM_STATUS == "C" | df$CLAIM_STATUS == "P"), "CLOSED", "OPEN")
    
    
#   Identify notification only claims that are not in Ocean data (df)
    noti_claims_not_in_Ocean <-  anti_join( x = noti_claims, y = df, by = "CLAIM_NUMBER") 
  
#   Combine all claims from Ocean (df) and notification only claims identified that were not in Ocean data (df). 
#   This is required to get the claims count correct. This should not have any effect on the financials - as there 
#   are no financial amounts associated with notification only claims.     
    df2 <- rbind(df,noti_claims_not_in_Ocean)
    

    
    
#     -------------------------------------------------------------------------------------------------  

    

    
      fins <- df2 %>%
        select(
            ACCIDENTYEAR,
            CLAIM_TYPE_NEW,
            MEDNETPAID,
            INDNETPAID,
            EXPNETPAID,
            MEDOPENRESERVE,
            INDOPENRESERVE,
            EXPOPENRESERVE,
            MEDNETINCURRED,
            INDNETINCURRED,
            EXPNETINCURRED
        ) %>%
        group_by(ACCIDENTYEAR, CLAIM_TYPE_NEW) %>%
        summarise(
            MED_PAID = sum(MEDNETPAID),
            IND_PAID = sum(INDNETPAID),
            EXP_PAID = sum(EXPNETPAID),
            MED_RESERVE = sum(MEDOPENRESERVE),
            IND_RESERVE = sum(INDOPENRESERVE),
            EXP_RESERVE = sum(EXPOPENRESERVE),
            MED_INCURRED = sum(MEDNETINCURRED),
            IND_INCURRED = sum(INDNETINCURRED),
            EXP_INCURRED = sum(EXPNETINCURRED)
        ) %>%
        mutate(DEVELOPMENT_MO = elapsed_months(as.Date(dates[dates.index], "%m/%d/%Y"),
                                               as.Date(paste0("01/01/", ACCIDENTYEAR), "%m/%d/%Y")
                                               )
               )
    
    
#   Combine with counts    
    
    
    fins_n_claims <- df2 %>%
        select(ACCIDENTYEAR,
               CLAIM_TYPE_NEW,
               CLAIM_STATUS_NEW) %>%
        group_by(ACCIDENTYEAR, CLAIM_TYPE_NEW, CLAIM_STATUS_NEW) %>%
        summarise(NO_CLAIMS = n()) %>%
        
        # filter(CLAIM_TYPE_NEW  == "MED_ONLY" ) %>%
        spread(CLAIM_STATUS_NEW, NO_CLAIMS) %>%
        #       Replaces NA's with 0 and adds Open and Closed claims.
        mutate(OPEN = replace_na(OPEN, 0),
               CLOSED = replace_na(CLOSED, 0),
               TOTAL_CLAIMS = (CLOSED + OPEN)) %>%
        left_join(fins, by = c("ACCIDENTYEAR",  "CLAIM_TYPE_NEW"))
    
  
    return(fins_n_claims)
    
     }    

 #     -------------------------------------------------------------------------------------------------      

    #   Initilize dfs object to store read-in files
    dfs <- NULL
    
    t0 <- Sys.time()
    
    #   Loop through the file.list and read them in one by one - and bind them into dfs object.
    for (i in 1:52) {
        gc()
        print(paste0("Please Wait...Now Importing Workbook -> '", file.list[i],"' and combining it with notification only claims valued as of: ", dates[i]))
        df <- create_df(file.list.index = i, dates.index = i)
        df$ACCIDENTYEAR <- as.numeric(df$ACCIDENTYEAR) # some years have it as char and some as num
        df$valued_as_of <- dates[i]
        dfs <- rbind(dfs,df)
        if (i == 52)
            cat(paste0(" Import Complete!\n "))
    }
    
    t1 <- Sys.time()
    t1-t0 
    
    #   Export results to .csv (Very Fast!)
    data.table::fwrite(dfs, file = "claims2.csv", row.names = FALSE)
   
    

    yr_end_tr_financial <- dfs[dfs$valued_as_of %in% yr_end_dt,]
    
    
    

    file.list[1]
    file.list[52]
    
    dates[1]
    dates[52]
