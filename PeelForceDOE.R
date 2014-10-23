# Load all librarys and load my custom functions
source(file="E:/Projects/Rutility/UtilityFunctions.R")
library(car)


#Read all the tabs of the excel data file
dataFileName <- tclvalue(tkgetOpenFile())
myData<-importWorksheets(dataFileName)
names(myData)

#select the worksheet from the object myData.
data<-myData$PeelForce
names(data)

#Identify the factors and fix some of the units on variables
data$SampleID <- factor(data$SampleID)
data$LidFilm <- factor(data$LidFilm)
data$Material <- factor(data$Material)
data$TimeS <- factor(data$TimeS)

tbl_df(data)

#Clean-up using tidyr - put all the measurements in 1 coulumn called Visual
#then spread them out one by one by parameter

mySummary <- data %>%
    select(LidFilm,Material,TimeS,TempC,ForcePpi) %>%
    spread(TempC,ForcePpi)

# Now save this summary in a spreadsheet
saveFileName <- tclvalue(tkgetSaveFile())
tabName="PeelForceTestSummary"
exportToExcel(mySummary,saveFileName,tabName)


# Build a plot using natural splines
plotTitle="Peel Force Effects Plot \n Cycle Time (s)"
xlabel="Sealing Temperature (deg C)"
ylabel="Peeling Force (lbf/in)"
p2 <- ggplot(data,aes(x=TempC,y=ForcePpi,group=LidFilm, colour=LidFilm, shape=LidFilm)) +
    geom_point(size=2, position=position_jitter(width=1.5, height=0.5)) +
    stat_smooth(method="lm",formula=y~ns(x,df=3), size=1, se=FALSE) +
    scale_y_continuous(minor_breaks = seq(0,4,0.5), breaks = seq(0,4,1)) +
    scale_x_continuous(minor_breaks = seq(110,200,5), breaks = seq(110,200,10)) +
    facet_grid(Material~TimeS) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme_bw()
ggsave(p2,width=11,height=6,file="PeelForceLinearEffects.png")
p2


#now fit the linear model and create the Anova table
mod.peel <- lm(ForcePpi ~ LidFilm*Material*ns(TempC, df=3)*TimeS,data=data)
summary(mod.peel)
Anova(mod.peel)
