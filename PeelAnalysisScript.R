# Load all librarys and load my custom functions
source(file="E:/Projects/Rutility/UtilityFunctions.R")
library(car)

#Read all the tabs of the excel data file
dataFileName <- tclvalue(tkgetOpenFile())
myData<-importWorksheets(dataFileName)
names(myData)

#select the worksheet from the object myData
data<-myData$FullBurstData
#data <- filter(data,TimeS==1)
names(data)

#Identify the factors and fix some of the units on variables
data$SampleID <- factor(data$SampleID)
data$LidFilm <- factor(data$LidFilm)
data$Material <- factor(data$Material)
data$TimeS <- factor(data$TimeS)

tbl_df(data)

#Clean-up using tidyr - put all the measurements in 1 coulumn called BurstHG
data <- data %>%
    gather(Tray,BurstHg,A:F)

# build a plot faceted by time and material
plotTitle="Burst Pressure Effects Plot \n Cycle Time (s)"
xlabel="Sealing Temperature (deg C)"
ylabel="Burst Pressure (in Hg)"
p1 <- ggplot(data,aes(x=TempC,y=BurstHg,group=LidFilm, colour=LidFilm, shape=LidFilm)) +
    geom_point(size=2, position=position_jitter(width=1.5, height=0.5)) +
    stat_smooth(span=0.4, size=1, se=FALSE) +
    scale_y_continuous(minor_breaks = seq(0,15,1), breaks = seq(0,15,5)) +
    scale_x_continuous(minor_breaks = seq(110,200,5), breaks = seq(110,200,10)) +
    facet_grid(Material~TimeS) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme_bw()
ggsave(p1,width=11,height=6,file="BurstPressureEffects.png")
p1

# Summarize the raw data in two steps: use plyr and tidyr
# 1st compute the mean BurstHg for each trial (six samples)
# 2nd spread the temperatures from one column to  individual columns and make a table
mySummary <- ddply(data,.(TimeS,LidFilm,Material,TempC),numcolwise(mean))
mySummary <- spread(mySummary,TempC,BurstHg)

# Now save this summary in a spreadsheet
saveFileName <- tclvalue(tkgetSaveFile())
tabName="BurstTestSummary"
exportToExcel(mySummary,saveFileName,tabName)


mod.burst <- lm(BurstHg ~ LidFilm+Material+TempC+TimeS,data=data)
summary(mod.burst)
Anova(mod.burst)
