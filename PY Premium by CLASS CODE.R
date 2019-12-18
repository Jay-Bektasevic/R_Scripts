####################################################################################################
#                                                                                                  #
# Purpose:       PY Premium by CLASS CODE                                                          #
#                                                                                                  #
# Author:        Jay Bektasevic                                                                    #
# Contact:       jbektasevic@clearpathmutual.com                                                   #
# Client:        Jay Bektasevic                                                                    #
#                                                                                                  #
# Code created:  2019-04-09                                                                        #
# Last updated:  2019-04-09                                                                        #
# Source:        K:/Departments/Accounting/Jay Bektasevic/Prem Forecast BREEZE                     #
#                                                                                                  #
# Comment:       This is ad hoc script to calculate PY premium by CLASS CODE. Which in turn will   #
#                be used to calculate Loss Ratio by Class Code.                                    #
#                                                                                                  #
####################################################################################################


####################################################################################################
#                                                                                                  #
#                        Set the environment and load the requied packages.                        #
#                                                                                                  #
####################################################################################################

#   Increase the heap size before the rJava package is loaded - it's a dependency for "xlsx" package.
    options(java.parameters = "-Xmx16000m")

#   List of packages to load:
    packages <- c('readxl',
                  'dplyr',
                  'xlsx',
                  "tidyr",
                  "RODBC")
# formatR::tidy_app()
    new_packages <-
        packages[!(packages %in% installed.packages()[, "Package"])]
    
    if (length(new_packages))
        install.packages(new_packages)

#   Load Neccessary Packages
    sapply(packages, require, character.only = TRUE)


    setwd("K:/Departments/Accounting/Jay Bektasevic/Book of Business Analysis/Class Code Analysis")

####################################################################################################
#                                                                                                  #
#                            Helper function - to clean up sql scripts.                            #
#                                                                                                  #
####################################################################################################

    doSub <- function(src,
                      dest_var_name,
                      src_pattern,
                      dest_pattern) {
        assign(
            x       = dest_var_name,
            value   = gsub(
                pattern = src_pattern,
                replacement = dest_pattern,
                x = src
            ),
            envir   = .GlobalEnv
        )
    }
    
################## Create new function to utilize the one above to clean up SQL  ###################
    
    clean_text <- function(fileName){
        
        original_text <- readChar(fileName, file.info(fileName)$size)
        
########################### Convert to UNIX line ending for ease of use ############################
        
        doSub(
            src = original_text,
            dest_var_name = 'unix_text',
            src_pattern = '\r\n',
            dest_pattern = '\n'
        )
        
###################################### Remove Block Comments #######################################
        doSub(
            src = unix_text,
            dest_var_name = 'wo_bc_text',
            src_pattern = '/\\*.*?\\*/',
            dest_pattern = ''
        )
        
####################################### Remove Line Comments #######################################
        doSub(
            src = wo_bc_text,
            dest_var_name = 'wo_bc_lc_text',
            src_pattern = '--.*?\n',
            dest_pattern = ''
        )
        
############################### Remove Line Endings to get Flat Text ###############################
        doSub(
            src = wo_bc_lc_text,
            dest_var_name = 'flat_text',
            src_pattern = '\n',
            dest_pattern = ' '
        )
        
##################################### Remove Contiguous Spaces #####################################
        doSub(
            src = flat_text,
            dest_var_name = 'clean_flat_text',
            src_pattern = ' +',
            dest_pattern = ' '
        )
    }
    
####################################################################################################
#                                                                                                  #
#                           Create Function to Iterate Through the Years                           #
#                                                                                                  #
####################################################################################################
    
    get_premiums <- function(policy_yr, as_of) {
        
#################################### Read in SQL File Contents #####################################
        
    pre_sql <- paste0("DECLARE @policy_year INT= ", policy_yr, "; ",
                      "DECLARE @AsOf DATE = '", as_of, "';"
                      )
    # policy_yr <- 2018
    # as_of <- "12/31/2018"
    fileName <- "Policy Year Premium Loss Ratio by Class Code.sql"

    clean_text(fileName)


####################################################################################################
#                                                                                                  #
#                            Create a channel to connect to SQL SERVER                             #
#                                                                                                  #
####################################################################################################

    con  <-  odbcConnect("RSQL_BI") # "RSQL" is a pre set-up DSN on my machine

########################################### RUN SQL CODE ###########################################

    clean_sql <- paste0(pre_sql, clean_flat_text)
    
    data <- sqlQuery(con, clean_sql)

####################################################################################################
#                                                                                                  #
# Here we will recreate CPM premium levels to make sure that our algorithm matches what we have    #
# in the system.  This section can be utilized for Rate making data submission.                    #
#                                                                                                  #
####################################################################################################


############################################# PAYROLL ##############################################

  #  sum(data$EXPOSURE_AMOUNT) 

########################################## MANUAL PREMIUM ##########################################

    data[data$EXPOSURE_AMOUNT == 0, which(colnames(data) == "EXPOSURE_AMOUNT")] <-  0.00000000001

#   These class Codes are fixed c(913, 908 )

    data$MANUAL_PREMIUM <- ifelse(
        data$CLASS_CODE %in% c(913, 908) ,
        data$EXPOSURE_AMOUNT * data$PAYROLL_RATE,
        (data$EXPOSURE_AMOUNT / 100) * data$PAYROLL_RATE
    )

######################################### MODIFIED PREMIUM #########################################

    data$EFFECTIVE_DATE <- as.Date(data$EFFECTIVE_DATE, "%m-%d-%Y")
    data$EXPIRATION_DATE <- as.Date(data$EXPIRATION_DATE, "%m-%d-%Y")

#   Identify jan and feb deductable policies because the deductable credit comes after SCHCR becaseu we did not adopt NCCI rates until 3/1/2018.
    jan_feb_ded_ins <- unique(data[data$EFFECTIVE_DATE < as.Date("03-01-2018", "%m-%d-%Y") & !is.na(data$DEDUCTIBLE_RATE), 3])

#   Fix the policies that have no mod to 1
    data[is.na(data$EMOD_RATE)  , which(colnames(data) == "EMOD_RATE")] <- 1


#   Please note that this calculation excludes deductable credit for deductable policies with effective date of < 3/1/2018.
#   For these policies - apply deductable credit after SCHCR to give you Standard Premium.

    data$MOD_PREM <- data$MANUAL_PREMIUM * (ifelse(
        !is.na(data$DEDUCTIBLE_RATE) &
            !data$INSURED_NUMBER %in% jan_feb_ded_ins,
        (1 - as.numeric(data$DEDUCTIBLE_RATE / 100)),
        1
    )) * (ifelse(
        !is.na(data$SHORT_RATE_PENALTY),
        as.numeric(data$SHORT_RATE_PENALTY),
        1
    ))  * (as.numeric(data$EMOD_RATE))


######################################### STANDARD PREMIUM #########################################

#   First, we need to get the total exposure of each polcy to use for minimum adjustment so to 
#   allocate correct amount of minimum adjustment for each class code.

    pol_exposure <- data %>%
        select(POLICY_ID,
               EXPOSURE_AMOUNT) %>%
        group_by(POLICY_ID) %>%
        summarise(pol_exposure = sum(EXPOSURE_AMOUNT))
    
    
    data <- data %>%
        left_join(pol_exposure, by = "POLICY_ID") %>%
        mutate(pct_of_total_exposure = (EXPOSURE_AMOUNT / pol_exposure))


    data$STANDARD_PREMIUM <-
        data$MOD_PREM * ifelse(!is.na(data$CREDIT_RATE),  (1 - as.numeric(data$CREDIT_RATE) / 100), 1) *
        
        ifelse(data$INSURED_NUMBER %in% jan_feb_ded_ins & !is.na(data$DEDUCTIBLE_RATE), (1 - as.numeric(data$DEDUCTIBLE_RATE / 100)), 1)  + 

#   Add minimum premium adjustment for policies that require it.
        ifelse(
            !is.na(data$MIN_PREM_ADJUSTMENT),
            (data$MIN_PREM_ADJUSTMENT * data$pct_of_total_exposure), 0
        )


########################################## POLICY PREMIUM ##########################################

    data$POLICY_PREMIUM <-
        data$STANDARD_PREMIUM * ifelse(!is.na(data$DISCOUNT_RATE),  (1 - as.numeric(data$DISCOUNT_RATE) / 100), 1) +
        
#   Apply terrorism Act and DTEC Act for every 100 of payroll
        ifelse(!is.na(data$TIA_RATE),
               (as.numeric(data$TIA_RATE) / 100) * (data$EXPOSURE_AMOUNT / 100),
               0) +
        ifelse(!is.na(data$DTC_RATE),
               (as.numeric(data$DTC_RATE) / 100) * (data$EXPOSURE_AMOUNT / 100),
               0)

    
    return(data)
    
    
    }  
 
   
####################################################################################################
#                                                                                                  #
#              Run get_premiums() function and bind results together in one dataframe              #
#                                                                                                  #
####################################################################################################
    
    prems <- NULL
    prems_final <- NULL
   
    t0 <- Sys.time()
    
    for (i in 2009:2019) {
        
        gc()
        print(paste0("Please Wait...Now Fetching ", i, " Premiums from BREEZE...Query run on: ", Sys.time()))
        prems <- get_premiums(policy_yr = i, as_of = "06/30/2019")
        prems_final <- rbind(prems_final, prems)
        if (i == 2019)
            cat(" Job Complete!\n ")
    }
    
    t1 <- Sys.time()
    t1 - t0
    
####################################################################################################
#                                                                                                  #
#                                       IMPORT CLAIMS DATA                                         #
#                                                                                                  #
####################################################################################################
    
    # Use this if you wnat to import from dataset already exported, otherwise use query to directly import 
    # from database.
    # claims <- data.table::fread(
    #     "K:/Departments/Accounting/Devi/Claims Data 2018 Year End/claims_2019-03-31T16-30-17.csv" 
    #     ) [POLICY_YEAR %in% c(2013:2017), c(2, 4, 5, 10, 11, 13, 58, 80, 81)]

    fileName <- "K:/Departments/Accounting/Jay Bektasevic/Financial Call Reporting/claims2013_2017.sql"
    
    clean_text(fileName)
    
    clean_sql <- clean_flat_text
    
    con  <-  odbcConnect("RSQL_BI") # "RSQL" is a pre set-up DSN on my machine
    
    claims <- sqlQuery(con, clean_sql)
    
    
####################################################################################################
#                                                                                                  #
#                                     IMPORT CLASS CODE DATA                                       #
#                                                                                                  #
####################################################################################################

    fileName <- "K:/Departments/Accounting/Jay Bektasevic/CC_to_NAICS.sql"

    clean_text(fileName)

    clean_sql <- clean_flat_text

    CLASS_CODE <- sqlQuery(con, clean_sql)

 #  Extract first 3 of NAICS code 
    CLASS_CODE$NAICS3 <- as.numeric(substr(CLASS_CODE$NAICS_CODE, 1, 3))
    
 #  Get descriptions for NAICS3   
    NAICS <- read_excel("C:/Users/jbektasevic/Downloads/2017_NAICS_Structure.xlsx", sheet = "NAICS3") 
    
    CLASS_CODE <- left_join(CLASS_CODE, NAICS, by = c("NAICS3"))
    
    

################################ GET LOSS VALUES BY CLASS CODE  ###################################   

  
    claims<-  claims %>%
    left_join(CLASS_CODE, by = c("CLASS_CODE") )
    
    
    clms <- claims %>%
        select(
            #POLICY_YEAR,
            NAICS3,
            CLAIM_NUMBER
        ) %>%
        group_by( 
            #POLICY_YEAR,
            NAICS3
            ) %>% 
        summarise(
            clms = n()
        )
    
    
############################# GET PREMIUM VALUES BY NAICS PREM SIZE #############################   
    
    
    pol_premium <- prems_final %>%
        select(POLICY_ID,
               POLICY_PREMIUM
               ) %>%
        group_by(POLICY_ID) %>%
        summarise(pol_premium = sum(POLICY_PREMIUM)) %>%
        mutate(
            pol_size = case_when(
                pol_premium < 1000 ~ "< 1,000",
                pol_premium >= 1000 & pol_premium < 5000 ~ "1,000 - 5,000",
                pol_premium >= 5000 & pol_premium < 10000 ~ "5,000 - 10,000",
                pol_premium >= 10000 & pol_premium < 25000 ~ "10,000 - 25,000",
                pol_premium >= 25000 & pol_premium < 50000 ~ "25,000 - 50,000",
                pol_premium >= 50000 & pol_premium < 100000 ~ "50,000 - 100,000",
                pol_premium > 100000 ~ "> 100,000"
            )
        )
    
  
    prems_final <- prems_final %>%
        left_join(pol_premium, by = "POLICY_ID") 
    
    
    
    size <- prems_final %>%
        select(
            POLICY_PREMIUM,
            pol_size
        ) %>%
        group_by(pol_size) %>%
        summarise(prem = sum(POLICY_PREMIUM) )
    
   
   
################################ GET LOSS VALUES BY POLICY   ######################################   
    
    clms <- claims %>%
        select(
            POLICY_YEAR,
            POLICY_ID,
            INCURRED_TOTAL
        ) %>%
        group_by( 
            POLICY_YEAR,
            POLICY_ID
        ) %>% 
        summarise(
            LOSS = sum(INCURRED_TOTAL),
            NO_CLMS = n()
        ) 
    
    
    clms <- claims %>%
        select(
            POLICY_YEAR,
            CLAIM_NUMBER,
            HAZARD_GROUP
        ) %>%
        group_by( 
            POLICY_YEAR,
            HAZARD_GROUP
        ) %>% 
        summarise(
            
            NO_CLMS = n_distinct(CLAIM_NUMBER)
        ) %>%
        spread(POLICY_YEAR, NO_CLMS)
    
#   Fix hazard group 3 to C    
    prems_final[prems_final$HAZARD_GROUP == 3, which(colnames(prems_final)== "HAZARD_GROUP")] <- "C"   
    
    premium <-   prems_final %>%
        filter(
            !is.na(CREDIT_RATE)
        ) %>%
        
        select(
            POLICY_YEAR,
            #POLICY_ID,
            #CLASS_CODE
            pol_size,
            CREDIT_RATE
            #HAZARD_GROUP
            
        ) %>%
        group_by(POLICY_YEAR,
                 pol_size
                 # CLASS_CODE.
                 #HAZARD_GROUP
        ) %>%
        summarise(
            
            #NO_POL = n_distinct(POLICY_ID)
            Avg_CR_DB = mean(CREDIT_RATE)
        )%>%
        spread(POLICY_YEAR, Avg_CR_DB) %>%
        
        arrange(factor(
            pol_size,
            levels = c('< 1,000',
                       '1,000 - 5,000',
                       '5,000 - 10,000',
                       '10,000 - 25,000',
                       '25,000 - 50,000',
                       '50,000 - 100,000',
                       '> 100,000')
        )
        )
    
    
    
    
     premium <-   prems_final %>%
        select(
            #POLICY_YEAR,
            POLICY_ID,
            #CLASS_CODE,
            pol_size,
            POLICY_PREMIUM
            
        ) %>%
        group_by(#POLICY_YEAR, 
            POLICY_ID,
                 pol_size
            #,CLASS_CODE
            ) %>%
        summarise(

            POLICY_PREMIUM =  round(sum(POLICY_PREMIUM), 0),
            NO_POL = n()
            
        ) %>%
        filter(POLICY_PREMIUM != 0) %>%
         left_join(clms,
                   by = c("POLICY_ID")) %>%
         
         mutate(
             LOSS = replace_na(LOSS, 0),
             NO_CLMS = replace_na(NO_CLMS, 0)
         ) %>%
         
         
        select (#POLICY_YEAR.x,
                POLICY_ID,
                pol_size,
                POLICY_PREMIUM,
                LOSS
            
        ) %>%
       
          group_by(#POLICY_YEAR.x, 
                 pol_size) %>%
        summarise(POLICY_PREMIUM = sum(POLICY_PREMIUM),
                  LOSS = sum(LOSS)) %>%
        
        mutate(LOSS_RATIO = LOSS / POLICY_PREMIUM) %>%
         
         arrange(factor(
             pol_size,
             levels = c('< 1,000',
                        '1,000 - 5,000',
                        '5,000 - 10,000',
                        '10,000 - 25,000',
                        '25,000 - 50,000',
                        '50,000 - 100,000',
                        '> 100,000')
         )
         )
     
     
         select(
             POLICY_YEAR.x, 
             pol_size,  CLM_COUNT
         ) %>%
         spread(POLICY_YEAR.x, CLM_COUNT) %>%
         
         arrange(factor(
             pol_size,
             levels = c('< 1,000',
                        '1,000 - 5,000',
                        '5,000 - 10,000',
                        '10,000 - 25,000',
                        '25,000 - 50,000',
                        '50,000 - 100,000',
                        '> 100,000')
                        )
         )
         
         
         
         rowwise()%>%
         mutate(
             AVG_5YR_LOSS_RATIO = mean(c(`2013`, `2014`, `2015`, `2016`, `2017`))
         )
       
    
    
################################ GET PREMIUM VALUES BY NAICS  #################################    
    
    premium <-   prems_final %>%
        select(
            #POLICY_YEAR,
            CLASS_CODE,
            EXPOSURE_AMOUNT,
            MANUAL_PREMIUM,
            MOD_PREM,
            STANDARD_PREMIUM,
            POLICY_PREMIUM
            
        ) %>%
        group_by(#POLICY_YEAR,
            CLASS_CODE) %>%
        summarise(
            PAYROLL = round(sum(EXPOSURE_AMOUNT), 0),
            MANUAL_PREMIUM = round(sum(MANUAL_PREMIUM), 0),
            MOD_PREMIUM = round(sum(MOD_PREM), 0),
            STD_PREMIUM = round(sum(STANDARD_PREMIUM), 0),
            POLICY_PREMIUM =  round(sum(POLICY_PREMIUM), 0)
            
        ) %>%
        filter(PAYROLL != 0) %>%
        left_join(CLASS_CODE,
                  by = "CLASS_CODE") %>%
        left_join(clms,
                  by = c("CLASS_CODE")) %>%
        mutate_if(is.numeric,
                  funs(replace_na(., 0))) %>%
        
        
        select (
            PAYROLL,
            POLICY_PREMIUM,
            #HAZARD_GROUP,
            NAICS_CODE3,
            BUSINESS_CATEGORY_DESCRIPTION,
            NAICS3_DESC,
            
            LOSS
        ) %>%
        group_by(NAICS_CODE3, BUSINESS_CATEGORY_DESCRIPTION, NAICS3_DESC) %>%
        summarise(premium = sum(POLICY_PREMIUM),
                  loss = sum(LOSS)) %>%
        
        mutate(LOSS_RATIO = loss / premium) %>%
        arrange(desc(LOSS_RATIO))
    
       
    write.xlsx(as.data.frame(premium), file = "class_code_loss_ratio.xlsx", 
               sheetName= "NAICS3", append = TRUE, row.names = FALSE, showNA = FALSE)  

    ################################ GET PREMIUM VALUES BY CLASS CODE  #################################   
    
    
    premium <-   prems_final %>%
        select(
            #POLICY_YEAR,
            CLASS_CODE,
            
            EXPOSURE_AMOUNT,
            MANUAL_PREMIUM,
            MOD_PREM,
            STANDARD_PREMIUM,
            POLICY_PREMIUM
            
        ) %>%
        group_by(#POLICY_YEAR,
            CLASS_CODE) %>%
        summarise(
            PAYROLL = round(sum(EXPOSURE_AMOUNT), 0),
            MANUAL_PREMIUM = round(sum(MANUAL_PREMIUM), 0),
            MOD_PREMIUM = round(sum(MOD_PREM), 0),
            STD_PREMIUM = round(sum(STANDARD_PREMIUM), 0),
            POLICY_PREMIUM =  round(sum(POLICY_PREMIUM), 0)
            
        ) %>%
        filter(PAYROLL != 0) %>%
        left_join(CLASS_CODE,
                  by = "CLASS_CODE") %>%
        left_join(clms,
                  by = c("CLASS_CODE")) %>%
        mutate_if(is.numeric,
                  funs(replace_na(., 0))) %>%
        
        
        
        select (
            PAYROLL,
            POLICY_PREMIUM,
            HAZARD_GROUP,
           CLASS_CODE,
           CC_DESCRIPTION,
            
            LOSS
        ) %>%
        group_by(CLASS_CODE, HAZARD_GROUP, CC_DESCRIPTION) %>%
        summarise(premium = sum(POLICY_PREMIUM),
                  loss = sum(LOSS)) %>%
        
        mutate(LOSS_RATIO = loss / premium) %>%
        arrange(desc(LOSS_RATIO))
    
    
    write.xlsx(as.data.frame(premium), file = "class_code_loss_ratio.xlsx", 
               sheetName= "CLASS_CODE", append = TRUE, row.names = FALSE, showNA = FALSE)  
    
    
    write.xlsx(as.data.frame(CLASS_CODE), file = "class_code_loss_ratio.xlsx", 
               sheetName= "Class Mapping", append = TRUE, row.names = FALSE, showNA = FALSE) 
    
    
    
#################################### Claims with Caped Losses  #####################################
   
    over250k_claims <- claims[claims$INCURRED_TOTAL > 250000,]
    
    z <- over250k_claims %>%
        select(
            NAICS3,
            CLAIM_NUMBER
        ) %>%
        group_by( NAICS3) %>%
        summarise(
            NO_CLM = n()
        )
    
    claims$INCURRED_TOTAL2 <- ifelse(claims$INCURRED_TOTAL > 250000, 250000, claims$INCURRED_TOTAL )
    
 x <-    claims %>%
        select(
            NAICS3,
            CLAIM_NUMBER,
            INCURRED_TOTAL,
            INCURRED_TOTAL2
        ) %>%
        group_by( NAICS3) %>%
        summarise(
            NO_CLM = n(),
            LOSS = sum(INCURRED_TOTAL),
            LOSS2 = sum(INCURRED_TOTAL2)
        ) %>%
     left_join(z, by = "NAICS3")%>%
     mutate_if(is.numeric,
               funs(replace_na(., 0)))

 
 
 z <-    claims %>%
     select(
         CLASS_CODE,
         POLICY_ID
         
     ) %>%
     group_by( CLASS_CODE) %>%
     summarise(
         NO_CLM = n()
         #LOSS = sum(INCURRED_TOTAL)
     ) %>% 
     mutate(
         AVG_CLM = LOSS / NO_CLM
     )
 
 
 prems_final <- prems_final %>%
     left_join(
         CLASS_CODE, by = "CLASS_CODE"
     )
 
 
 
 premium <-   prems_final %>%
     select(
         POLICY_YEAR,
         NAICS3,
         
         POLICY_ID
     ) %>%
     group_by(POLICY_YEAR,
              NAICS3) %>%
     summarise(
        NO_POL = n()
     ) %>%
     spread(
         POLICY_YEAR,
         NO_POL
     ) %>%
     mutate_if(is.numeric,
               funs(replace_na(., 0)))
     
 
 
 premium <-   prems_final %>%
     select(
         #POLICY_YEAR,
         CLASS_CODE,
         
         EXPOSURE_AMOUNT,
    
         STANDARD_PREMIUM
         
     ) %>%
     group_by(#POLICY_YEAR,
         CLASS_CODE) %>%
     summarise(
         PAYROLL = round(sum(EXPOSURE_AMOUNT), 0),

         STANDARD_PREMIUM =  round(sum(STANDARD_PREMIUM), 0)
         
     )

 
 
 
 
 write.xlsx(as.data.frame(premium), file = "class_code_loss_ratio.xlsx", 
            sheetName= "expo", append = TRUE, row.names = FALSE, showNA = FALSE) 

 
 ##################################### Depricated CLASS CODES  ######################################
     
 
 cc_depricated <- c('7228', '7229', '2747', '3069', '913', '7613', '4061', '4439', '908', '7611', '7423')
 
 cc_replaced <- c('7219', '7219', '2881', '3076', '913', '7600', '4062', '4558', '908', '7600', '7403')
 
 depricaded_ccs <- data.frame(cc_depricated, cc_replaced)
 
 
 ################################### AVG SCHD CD/DB BY CLASS CODE ################################### 
 
 cc_premium <- prems_final %>%
     select(CLASS_CODE,
            POLICY_PREMIUM
     ) %>%
     group_by(CLASS_CODE) %>%
     summarise(cc_premium = sum(POLICY_PREMIUM)) 
 
 
 x <- prems_final %>%
 
 left_join(cc_premium, by = "CLASS_CODE") 
 
 
 x <-   x %>%
     mutate(
         CREDIT_RATE = replace_na(CREDIT_RATE,0),
         scd_weighted = (POLICY_PREMIUM/cc_premium) * CREDIT_RATE
     ) %>%
     
     select(
         POLICY_YEAR,
         CLASS_CODE,

         CREDIT_RATE
         #,scd_weighted
         
     ) %>%
     group_by(
         POLICY_YEAR,
         CLASS_CODE
         ) %>%
     summarise(
        
         SCHD_CR = mean(CREDIT_RATE)
         #,W_SCHD_CR = sum(scd_weighted)
         
     ) %>%
     spread(
         POLICY_YEAR,
         SCHD_CR
         
     )
         
         


write.xlsx(as.data.frame(x), file = "class_code_loss_ratio.xlsx", 
           sheetName= "SCD", append = TRUE, row.names = FALSE, showNA = FALSE) 



x <-   x %>%
select(
    POLICY_YEAR,
   
    
    CREDIT_RATE
    #,scd_weighted
    
) %>%
    group_by(
        POLICY_YEAR
    ) %>%
    summarise(
        
        SCHD_CR = mean(CREDIT_RATE)
        #,W_SCHD_CR = sum(scd_weighted)
        
    ) %>%
    spread(
        POLICY_YEAR,
        SCHD_CR
        
    ) 



over250k_claims %>% 
    select(
        HAZARD_GROUP.x,
        INCURRED_TOTAL
    ) %>% 
    group_by(
        HAZARD_GROUP.x
    ) %>%
    summarise(
        loss = sum(INCURRED_TOTAL)
    ) %>%
    mutate(
        pct = loss/ sum(loss)
    )
###########################################################

size_premium <- prems_final %>%
    select(pol_size,
           POLICY_PREMIUM
    ) %>%
    group_by(pol_size) %>%
    summarise(size_premium = sum(POLICY_PREMIUM)) 


x <- x %>%
    
    left_join(size_premium, by = "pol_size") 




x1 <-   x %>%
    filter(
        !is.na(CREDIT_RATE)
    )%>%
    mutate(
        #CREDIT_RATE = replace_na(CREDIT_RATE,0),
        scd_weighted = (POLICY_PREMIUM/size_premium) * CREDIT_RATE
    ) %>%
    
    select(
        POLICY_YEAR,
        #POLICY_ID,
        #CLASS_CODE
        pol_size,
        #CREDIT_RATE
        scd_weighted
        #HAZARD_GROUP
        
    ) %>%
    group_by(POLICY_YEAR,
             pol_size
             # CLASS_CODE.
             #HAZARD_GROUP
    ) %>%
    summarise(
        
        #NO_POL = n_distinct(POLICY_ID)
        w_Avg_CR_DB = sum(scd_weighted)
    )%>%
    spread(POLICY_YEAR, w_Avg_CR_DB) %>%
    
    arrange(factor(
        pol_size,
        levels = c('< 1,000',
                   '1,000 - 5,000',
                   '5,000 - 10,000',
                   '10,000 - 25,000',
                   '25,000 - 50,000',
                   '50,000 - 100,000',
                   '> 100,000')
    )
    )

