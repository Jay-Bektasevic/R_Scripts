#   It's no secret that Google Big Brothers most of us. 
#   But at least they allow us to access quite a lot of the data they have collected on us. 
#   Among this is the Google location history.
#   Go to your google account and download Location History.


library(jsonlite)
library(lubridate)
library(zoo)
library(ggplot2)
library(ggmap)
library(raster)

system.time(x <- fromJSON("Location History.json"))

# extracting the locations dataframe
loc = x$locations

# converting time column from posix milliseconds into a readable time scale
loc$time = as.POSIXct(as.numeric(x$locations$timestampMs)/1000, origin = "1970-01-01")

# converting longitude and latitude from E7 to GPS coordinates
loc$lat = loc$latitudeE7 / 1e7
loc$lon = loc$longitudeE7 / 1e7

head(loc)

# how many rows are in the data frame?
nrow(loc)
min(loc$time)
max(loc$time)

# calculate the number of data points per day, month and year
loc$date <- as.Date(loc$time, '%Y/%m/%d')
loc$year <- year(loc$date)
loc$month_year <- as.yearmon(loc$date)

points_p_day <- data.frame(table(loc$date), group = "day")
points_p_month <- data.frame(table(loc$month_year), group = "month")
points_p_year <- data.frame(table(loc$year), group = "year")

#How many days were recorded?
nrow(points_p_day)
#How many months?
nrow(points_p_month)
#And how many years?
nrow(points_p_year)
# set up plotting theme

my_theme <- function(base_size = 12, base_family = "sans"){
    theme_grey(base_size = base_size, base_family = base_family) +
        theme(
            axis.text = element_text(size = 12),
            axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1),
            axis.title = element_text(size = 14),
            panel.grid.major = element_line(color = "grey"),
            panel.grid.minor = element_blank(),
            panel.background = element_rect(fill = "aliceblue"),
            strip.background = element_rect(fill = "lightgrey", color = "grey", size = 1),
            strip.text = element_text(face = "bold", size = 12, color = "navy"),
            legend.position = "right",
            legend.background = element_blank(),
            panel.margin = unit(.5, "lines"),
            panel.border = element_rect(color = "grey", fill = NA, size = 0.5)
        )
}



points <- rbind(points_p_day[, -1], points_p_month[, -1], points_p_year[, -1])

ggplot(points, aes(x = group, y = Freq)) + 
    geom_point(position = position_jitter(width = 0.2), alpha = 0.3) + 
    geom_boxplot(aes(color = group), size = 1, outlier.colour = NA) + 
    facet_grid(group ~ ., scales = "free") + my_theme() +
    theme(
        legend.position = "none",
        #strip.placement = "outside",
        strip.background = element_blank(),
        strip.text = element_blank(),
        axis.text.x = element_text(angle = 0, vjust = 0.5, hjust = 0.5)
    ) +
    labs(
        x = "",
        y = "Number of data points",
        title = "How many data points did Google collect about me?",
        subtitle = "Number of data points per day, month and year",
        caption = "\nGoogle collected between 0 and 1500 data points per day
        (median ~500), between 0 and 40,000 per month (median ~15,000) and 
        between 80,000 and 220,000 per year (median ~140,000)."
    )

#How accurate is the data?
#Accuracy is given in meters, i.e. the smaller the better.

accuracy <- data.frame(accuracy = loc$accuracy, group = ifelse(loc$accuracy < 800, "high", ifelse(loc$accuracy < 5000, "middle", "low")))

accuracy$group <- factor(accuracy$group, levels = c("high", "middle", "low"))

ggplot(accuracy, aes(x = accuracy, fill = group)) + 
    geom_histogram() + 
    facet_grid(group ~ ., scales="free") + 
    my_theme() +
    theme(
        legend.position = "none",
        #strip.placement = "outside",
        strip.background = element_blank(),
        axis.text.x = element_text(angle = 0, vjust = 0.5, hjust = 0.5)
    ) +
    labs(
        x = "Accuracy in metres",
        y = "Count",
        title = "How accurate is the location data?",
        subtitle = "Histogram of accuracy of location points",
        caption = "\nMost data points are pretty accurate, 
        but there are still many data points with a high inaccuracy.
        These were probably from areas with bad satellite reception."
    )
#Plotting data points on maps
#Finally, we are actually going to plot some maps!
#The first map is a simple point plot of all locations recorded around Germany.
#You first specify the map area and the zoom factor (the furthest away is 1); the bigger the zoom, the closer to the center of the specified location. Location can be given as longitude/latitude pair or as city or country name.
#On this map we can plot different types of plot with the regular ggplot2 syntax. For example, a point plot.

USA <- get_map(location = 'United States', zoom = 5)

ggmap(USA) + geom_point(data = loc, aes(x = lon, y = lat), alpha = 0.5, color = "red") + 
    theme(legend.position = "right") + 
    labs(
        x = "Longitude", 
        y = "Latitude", 
        title = "Location history data points in US",
        caption = "\nA simple point plot shows recorded positions.")

#The second map shows a 2D bin plot of accuracy measured for all data points recorded in my home town Louisville.

louisville <- get_map(location = 'Louisville', zoom = 12)

options(stringsAsFactors = T)
ggmap(louisville) + 
    stat_summary_2d(geom = "tile", bins = 100, data = loc, aes(x = lon, y = lat, z = accuracy), alpha = 0.5) + 
    scale_fill_gradient(low = "blue", high = "red", guide = guide_legend(title = "Accuracy")) +
    labs(
        x = "Longitude", 
        y = "Latitude", 
        title = "Location history data points around Louisville",
        subtitle = "Color scale shows accuracy (low: blue, high: red)",
        caption = "\nThis bin plot shows recorded positions 
        and their accuracy in and around Louisville")


#We can also plot the velocity of each data point:

loc_2 <- loc[which(!is.na(loc$velocity)), ]

munster <- get_map(location = 'Munster', zoom = 10)

ggmap(louisville) + geom_point(data = loc_2, aes(x = lon, y = lat, color = velocity), alpha = 0.3) + 
    theme(legend.position = "right") + 
    labs(x = "Longitude", y = "Latitude", 
         title = "Location history data points in Louisville",
         subtitle = "Color scale shows velocity measured for location",
         caption = "\nA point plot where points are colored according 
         to velocity nicely reflects that I moved generally 
         slower in the city center than on the autobahn") +
    scale_colour_gradient(low = "blue", high = "red", guide = guide_legend(title = "Velocity"))

#What distance did I travel?
#To obtain the distance I moved, I am calculating the distance between data points. Because it takes a long time to calculate, I am only looking at data from last year.

loc3 <- with(loc, subset(loc, loc$time > as.POSIXct('2016-01-01 0:00:01')))
loc3 <- with(loc, subset(loc3, loc$time < as.POSIXct('2016-12-22 23:59:59')))

# Shifting vectors for latitude and longitude to include end position
shift.vec <- function(vec, shift){
    if (length(vec) <= abs(shift)){
        rep(NA ,length(vec))
    } else {
        if (shift >= 0) {
            c(rep(NA, shift), vec[1:(length(vec) - shift)]) }
        else {
            c(vec[(abs(shift) + 1):length(vec)], rep(NA, abs(shift)))
        }
    }
}

loc3$lat.p1 <- shift.vec(loc3$lat, -1)
loc3$lon.p1 <- shift.vec(loc3$lon, -1)

# Calculating distances between points (in metres) with the function pointDistance from the 'raster' package.

loc3$dist.to.prev <- apply(loc3, 1, FUN = function(row) {
    pointDistance(c(as.numeric(as.character(row["lat.p1"])),
                    as.numeric(as.character(row["lon.p1"]))),
                  c(as.numeric(as.character(row["lat"])), as.numeric(as.character(row["lon"]))),
                  lonlat = T) # Parameter 'lonlat' has to be TRUE!
})
# distance in km
round(sum(as.numeric(as.character(loc3$dist.to.prev)), na.rm = TRUE)*0.001, digits = 2)
## [1] 54466.08
distance_p_month <- aggregate(loc3$dist.to.prev, by = list(month_year = as.factor(loc3$month_year)), FUN = sum)
distance_p_month$x <- distance_p_month$x*0.001
ggplot(distance_p_month[-1, ], aes(x = month_year, y = x,  fill = month_year)) + 
    geom_bar(stat = "identity")  + 
    guides(fill = FALSE) +
    my_theme() +
    labs(
        x = "",
        y = "Distance in km",
        title = "Distance traveled per month in 2016",
        caption = "This barplot shows the sum of distances between recorded 
        positions for 2016. In October we went to Costa Rica."
    )
#Activities
#Google also guesses my activity based on distance travelled per time.
#Here again, it would take too long to look at activity from all data points.

activities <- loc3$activitys

list.condition <- sapply(activities, function(x) !is.null(x[[1]]))
activities  <- activities[list.condition]

df <- do.call("rbind", activities)
main_activity <- sapply(df$activities, function(x) x[[1]][1][[1]][1])

activities_2 <- data.frame(main_activity = main_activity, 
                           time = as.POSIXct(as.numeric(df$timestampMs)/1000, origin = "1970-01-01"))

head(activities_2)

ggplot(activities_2, aes(x = main_activity, group = main_activity, fill = main_activity)) + 
    geom_bar()  + 
    guides(fill = FALSE) +
    my_theme() +
    labs(
        x = "",
        y = "Count",
        title = "Main activities in 2016",
        caption = "Associated activity for recorded positions in 2016. 
        Because Google records activity probabilities for each position, 
        only the activity with highest likelihood were chosen for each position."
    )


# Since, google now requires API to map it, I will just export it and view it in Tableau.

x <- as.data.frame(loc[,10:13])
write.csv(x, file = "locationHistory.csv")
