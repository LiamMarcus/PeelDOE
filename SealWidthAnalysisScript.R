# Load all librarys and load my custom functions
source(file="E:/Projects/Rutility/UtilityFunctions.R")
library(car)

#Read all the tabs of the excel data file
dataFileName <- tclvalue(tkgetOpenFile())
myData<-importWorksheets(dataFileName)
names(myData)

#select the worksheet from the object myData.
data<-myData$SealWidthData
names(data)

#Identify the factors and fix some of the units on variables
data$SampleID <- factor(data$SampleID)
data$LidFilm <- factor(data$LidFilm)
data$Material <- factor(data$Material)
data$TimeS <- factor(data$TimeS)

tbl_df(data)

# Summarize the raw data in two steps: use plyr and tidyr
# 1st compute the mean SealWidth for each trial (three samples)
# 2nd spread the temperatures from one column to  individual columns and make a table
mySummary <- ddply(data,.(TimeS,LidFilm,Material,TempC),numcolwise(mean))
mySummary <- spread(mySummary,TempC,SealWidth)

# Now save this summary in a spreadsheet
saveFileName <- tclvalue(tkgetSaveFile())
tabName="SealWidthSummary"
exportToExcel(mySummary,saveFileName,tabName)

# build a plot faceted by time and material
plotTitle="Seal Width Effects Plot \n Cycle Time (s)"
xlabel="Sealing Temperature (deg C)"
ylabel="Seal Width (in)"
p1 <- ggplot(data,aes(x=TempC,y=SealWidth,group=LidFilm, colour=LidFilm, shape=LidFilm)) +
    geom_point(size=2) +
#    stat_smooth(span=0.6, size=1, se=FALSE) +
    stat_smooth(method="lm",formula=y~ns(x,df=6), size=1, se=FALSE) +
    scale_y_continuous(minor_breaks = seq(0,0.25,0.01), breaks = seq(0,0.25,0.05)) +
    scale_x_continuous(minor_breaks = seq(110,200,5), breaks = seq(110,200,10)) +
    facet_grid(Material~TimeS) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme_bw()
ggsave(p1,width=11,height=6,file="SealWidthEffects.png")
p1
 



mod.burst <- lm(BurstHg ~ LidFilm+Material+TempC+TimeS,data=data)
summary(mod.burst)
Anova(mod.burst)
