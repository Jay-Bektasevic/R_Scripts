
#   Increase the heap size before the rJava package is loaded - it's a dependency for "xlsx" package.
options(java.parameters = "-Xmx16000m")

#   List of packages to load:
packages <- c('readxl',
              'dplyr',
              "tidyr",
              "rjson"
)

new_packages <-
    packages[!(packages %in% installed.packages()[, "Package"])]

if (length(new_packages))
    install.packages(new_packages)

#   Load Neccessary Packages
sapply(packages, require, character.only = TRUE)





#   ------------------------------------------------------------------------------------
#   P_Ball drawing

# Get the history winnings from the URL.
json_file <- "https://data.ny.gov/resource/d6yy-54nr.json"

win_his <-
    dplyr::bind_rows(fromJSON(file = json_file)) %>% as.data.frame() %>%
    separate(winning_numbers, into = paste("B", 1:6, sep = "_")) %>%
    select(-multiplier) %>%
    mutate(draw_date = as.Date(draw_date))


#   Convert to long format
all_balls <-
    win_his %>% gather(key = 'MEASURE', value = 'VALUE', B_1:B_6)

#   Last time the balls were drawn
white_last_draw <- all_balls %>%
    filter(MEASURE != "B_6",
           draw_date > as.Date("10/07/2015", "%m/%d/%Y")) %>%
    arrange(draw_date) %>%
    group_by(VALUE) %>%
    slice(which.max(draw_date)) %>%
    mutate(days_since_draw = Sys.Date() - draw_date) %>%
    arrange(desc(days_since_draw))


red_last_draw <- all_balls %>%
    filter(MEASURE == "B_6",
           draw_date > as.Date("10/07/2015", "%m/%d/%Y")) %>%
    arrange(draw_date) %>%
    group_by(VALUE) %>%
    slice(which.max(draw_date)) %>%
    mutate(days_since_draw = Sys.Date() - draw_date) %>%
    arrange(desc(days_since_draw))


#   Top numbers drawn since oct 2015
top_10_white <- all_balls %>%
    filter(MEASURE != "B_6",
           draw_date > as.Date("10/07/2015", "%m/%d/%Y")) %>%
    group_by(VALUE) %>%
    tally() %>%
    ungroup %>%
    top_n(20) %>%
    mutate(
        pdf = n/nrow(all_balls[all_balls$MEASURE != "B_6" & all_balls$draw_date > as.Date("10/07/2015", "%m/%d/%Y") ,])
    ) %>%
    left_join(select(white_last_draw, VALUE, days_since_draw), by = c("VALUE" = "VALUE")) %>%
    arrange(desc(days_since_draw))

top_10_red <- all_balls %>%
    filter(MEASURE == "B_6",
           draw_date > as.Date("10/07/2015", "%m/%d/%Y")) %>%
    group_by(VALUE) %>%
    tally() %>%
    ungroup %>%
    top_n(10) %>%
    mutate(
        pdf = n/nrow(all_balls[all_balls$MEASURE == "B_6" & all_balls$draw_date > as.Date("10/07/2015", "%m/%d/%Y") ,])
    ) %>%
    left_join(select(red_last_draw, VALUE, days_since_draw), by = c("VALUE" = "VALUE")) %>%
    arrange(desc(days_since_draw)) 


  
 # get avg number of days between drawings of same number 
  
  x <- all_balls %>%
    filter(MEASURE == "B_6",
           draw_date > as.Date("10/07/2015", "%m/%d/%Y")) %>%
    left_join(select(top_10_red , VALUE, days_since_draw), by = c("VALUE" = "VALUE")) %>%
    filter(
        !is.na(days_since_draw)
    )
 z <- NULL
  
 for (i in 1:nrow(top_10_red)) {
   b <-  x %>% filter(
         VALUE == top_10_red$VALUE[i]
     ) %>%
         arrange(
             desc(draw_date)
         ) %>%
         
         mutate(Between = as.numeric(c(diff(draw_date),0)),
        avg_days_between = mean(abs(Between)),
        std = sd(abs(Between)),
        z_score = (days_since_draw - avg_days_between)/std
         ) 
   z <- rbind(z, b)
 }
   
top_10_red <- top_10_red %>% 
  left_join(select(z, VALUE, z_score),  by = c("VALUE" = "VALUE")  ) %>%
    group_by(VALUE) %>%
     filter(row_number() == 1) %>%
     mutate(
         new_prob = pdf * abs(z_score)
     )

 
 
   
  x <- all_balls %>%
    filter(MEASURE != "B_6",
           draw_date > as.Date("10/07/2015", "%m/%d/%Y")) %>%
    left_join(select(top_10_white , VALUE, days_since_draw), by = c("VALUE" = "VALUE")) %>%
    filter(
        !is.na(days_since_draw)
    )
 z <- NULL
  
 for (i in 1:nrow(top_10_white)) {
   b <-  x %>% filter(
         VALUE == top_10_white$VALUE[i]
     ) %>%
         arrange(
             desc(draw_date)
         ) %>%
         
         mutate(Between = as.numeric(c(diff(draw_date),0)),
        avg_days_between = mean(abs(Between)),
        std = sd(abs(Between)),
        z_score = (days_since_draw - avg_days_between)/std
         ) 
   z <- rbind(z, b)
 }
   
 top_10_white <- top_10_white %>% 
  left_join(select(z, VALUE, z_score),  by = c("VALUE" = "VALUE")  ) %>%
    group_by(VALUE) %>%
     filter(row_number() == 1) %>%
     mutate(
         new_prob = pdf * abs(z_score)
     )
 
 
# simulate
 
 
 df <- data.frame(matrix(ncol = 6, nrow = 0))
x <- c('B1',	'B2',	'B3',	'B4',	'B5',	'R6')
colnames(df) <- x

for (i in 1: 1000) {
  
df[i, 1:5] <- sample( t(top_10_white[,1]) , 5, replace = FALSE, prob = top_10_white$new_prob) 
df[i, 6] <- sample( t(top_10_red[,1]) , 1, replace = FALSE, prob = top_10_red$new_prob) 
}   


#   White Balls
  df1 <-  df %>% gather(key = 'MEASURE', value = 'VALUE', B1:R6) %>%
        filter(
            MEASURE != "R6"
        ) %>%
        group_by(VALUE) %>%
        tally() %>%
        top_n(20) %>%
        arrange(desc(n)) 
 

    

#   Red Ball    
df2 <-  df %>% gather(key = 'MEASURE', value = 'VALUE', B1:R6) %>%
        filter(
            MEASURE == "R6"
        ) %>%
        group_by(VALUE) %>%
        tally() %>%
        top_n(10) %>%
        arrange(desc(n))



# print numbers
cat(sample(t(df1[,1]), 5, replace = FALSE ),   sample(t(df2[,1]), 1, replace = FALSE ))
cat(sample(t(df1[,1]), 5, replace = FALSE ),   sample(t(df2[,1]), 1, replace = FALSE ))
cat(sample(t(df1[,1]), 5, replace = FALSE ),   sample(t(df2[,1]), 1, replace = FALSE ))

     
