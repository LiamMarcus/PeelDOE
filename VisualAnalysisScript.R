# Load all librarys and load my custom functions
source(file="E:/Projects/Rutility/UtilityFunctions.R")
library(car)

#Read all the tabs of the excel data file
dataFileName <- tclvalue(tkgetOpenFile())
myData<-importWorksheets(dataFileName)
names(myData)

#select the worksheet from the object myData.
data<-myData$SealAppearance
#data <- filter(data,TimeS==1)
names(data)

#Identify the factors and fix some of the units on variables
data$SampleID <- factor(data$SampleID)
data$LidFilm <- factor(data$LidFilm)
data$Material <- factor(data$Material)
data$TimeS <- factor(data$TimeS)

tbl_df(data)

# Summarize the raw data in two steps: use plyr and tidyr
# 1st compute the mean for each trial
# 2nd spread the temperatures from one column to  individual columns and make a table
mySummary <- ddply(data,.(TimeS,LidFilm,Material,TempC),numcolwise(mean))

#Clean-up using tidyr - put all the measurements in 1 coulumn called Visual
#then spread them out one by one by parameter

myAppears <- mySummary %>%
    gather(Visual,Score,Appears:Caulking) %>%
    filter(Visual=="Appears") %>%
    spread(TempC,Score)

myStrings <- mySummary %>%
    gather(Visual,Score,Appears:Caulking) %>%
    filter(Visual=="Strings") %>%
    spread(TempC,Score)

myGaps <- mySummary %>%
    gather(Visual,Score,Appears:Caulking) %>%
    filter(Visual=="Gaps") %>%
    spread(TempC,Score)

myHairlines <- mySummary %>%
    gather(Visual,Score,Appears:Caulking) %>%
    filter(Visual=="Hairlines") %>%
    spread(TempC,Score)

myBubbles <- mySummary %>%
    gather(Visual,Score,Appears:Caulking) %>%
    filter(Visual=="Bubbles") %>%
    spread(TempC,Score)

myCaulking <- mySummary %>%
    gather(Visual,Score,Appears:Caulking) %>%
    filter(Visual=="Caulking") %>%
    spread(TempC,Score)

#now concatenate all these dataframes into one dataframe
visualSummary <- rbind(myAppears, myStrings, myGaps, myHairlines, myBubbles, myCaulking)

# Now save these summaries in a spreadsheet
saveFileName <- tclvalue(tkgetSaveFile())
tabName="VisualSummary"
exportToExcel(visualSummary,saveFileName,tabName)


# build a plot faceted by time and material
plotTitle="Visual Appearance Effects Plot \n Cycle Time (s)"
xlabel="Sealing Temperature (deg C)"
ylabel="Visual Appearance Score"
p1 <- ggplot(data,aes(x=TempC,y=Appears,group=LidFilm, colour=LidFilm, shape=LidFilm)) +
    geom_point(size=2, position=position_jitter(width=1.5, height=0.5)) +
#    stat_smooth(span=0.6, size=1, se=FALSE) +
    stat_smooth(method="lm",formula=y~ns(x,df=4), size=1, se=FALSE) +
    scale_y_continuous(minor_breaks = seq(0,6,0.5), breaks = seq(0,6,1)) +
    scale_x_continuous(minor_breaks = seq(110,200,5), breaks = seq(110,200,10)) +
    facet_grid(Material~TimeS) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme_bw()
ggsave(p1,width=11,height=6,file="AppearanceEffects.png")
p1

plotTitle="Stringy Appearance Effects Plot \n Cycle Time (s)"
xlabel="Sealing Temperature (deg C)"
ylabel="Visual Strings Score"
p1 <- ggplot(data,aes(x=TempC,y=Strings,group=LidFilm, colour=LidFilm, shape=LidFilm)) +
    geom_point(size=2, position=position_jitter(width=1.5, height=0.5)) +
#    stat_smooth(span=0.6, size=1, se=FALSE) +
    stat_smooth(method="lm",formula=y~ns(x,df=4), size=1, se=FALSE) +
    scale_y_continuous(minor_breaks = seq(0,6,0.5), breaks = seq(0,6,1)) +
    scale_x_continuous(minor_breaks = seq(110,200,5), breaks = seq(110,200,10)) +
    facet_grid(Material~TimeS) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme_bw()
ggsave(p1,width=11,height=6,file="StringEffects.png")
p1

plotTitle="Gap Appearance Effects Plot \n Cycle Time (s)"
xlabel="Sealing Temperature (deg C)"
ylabel="Visual Gaps Score"
p1 <- ggplot(data,aes(x=TempC,y=Gaps,group=LidFilm, colour=LidFilm, shape=LidFilm)) +
    geom_point(size=2, position=position_jitter(width=1.5, height=0.5)) +
#    stat_smooth(span=0.6, size=1, se=FALSE) +
    stat_smooth(method="lm",formula=y~ns(x,df=4), size=1, se=FALSE) +
    scale_y_continuous(minor_breaks = seq(0,6,0.5), breaks = seq(0,6,1)) +
    scale_x_continuous(minor_breaks = seq(110,200,5), breaks = seq(110,200,10)) +
    facet_grid(Material~TimeS) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme_bw()
ggsave(p1,width=11,height=6,file="GapEffects.png")
p1

plotTitle="Hairline Appearance Effects Plot \n Cycle Time (s)"
xlabel="Sealing Temperature (deg C)"
ylabel="Visual Hairlines Score"
p1 <- ggplot(data,aes(x=TempC,y=Hairlines,group=LidFilm, colour=LidFilm, shape=LidFilm)) +
    geom_point(size=2, position=position_jitter(width=1.5, height=0.5)) +
#    stat_smooth(span=0.6, size=1, se=FALSE) +
    stat_smooth(method="lm",formula=y~ns(x,df=4), size=1, se=FALSE) +
    scale_y_continuous(minor_breaks = seq(0,6,0.5), breaks = seq(0,6,1)) +
    scale_x_continuous(minor_breaks = seq(110,200,5), breaks = seq(110,200,10)) +
    facet_grid(Material~TimeS) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme_bw()
ggsave(p1,width=11,height=6,file="HairlineEffects.png")
p1

plotTitle="Bubble Appearance Effects Plot \n Cycle Time (s)"
xlabel="Sealing Temperature (deg C)"
ylabel="Visual Bubbles Score"
p1 <- ggplot(data,aes(x=TempC,y=Bubbles,group=LidFilm, colour=LidFilm, shape=LidFilm)) +
    geom_point(size=2, position=position_jitter(width=1.5, height=0.5)) +
#    stat_smooth(span=0.6, size=1, se=FALSE) +
    stat_smooth(method="lm",formula=y~ns(x,df=4), size=1, se=FALSE) +
    scale_y_continuous(minor_breaks = seq(0,6,0.5), breaks = seq(0,6,1)) +
    scale_x_continuous(minor_breaks = seq(110,200,5), breaks = seq(110,200,10)) +
    facet_grid(Material~TimeS) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme_bw()
ggsave(p1,width=11,height=6,file="BubbleEffects.png")
p1

plotTitle="Caulking Appearance Effects Plot \n Cycle Time (s)"
xlabel="Sealing Temperature (deg C)"
ylabel="Visual Caulking Score"
p1 <- ggplot(data,aes(x=TempC,y=Caulking,group=LidFilm, colour=LidFilm, shape=LidFilm)) +
    geom_point(size=2, position=position_jitter(width=1.5, height=0.5)) +
#    stat_smooth(span=0.6, size=1, se=FALSE) +
    stat_smooth(method="lm",formula=y~ns(x,df=4), size=1, se=FALSE) +
    scale_y_continuous(minor_breaks = seq(0,6,0.5), breaks = seq(0,6,1)) +
    scale_x_continuous(minor_breaks = seq(110,200,5), breaks = seq(110,200,10)) +
    facet_grid(Material~TimeS) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme_bw()
ggsave(p1,width=11,height=6,file="CaulkingEffects.png")
p1
