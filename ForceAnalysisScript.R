# Load all librarys and load my custom functions
source(file="E:/Projects/Rutility/UtilityFunctions.R")

#Read all the tabs of the excel data file
dataFileName <- tclvalue(tkgetOpenFile())
myData<-read.csv(dataFileName)
names(myData)

#select the worksheet from the object myData.
tbl_df(myData)

data <- filter(myData,TimeS>=10 & TimeS<=20 | TimeS >=50 & TimeS <=60)
tbl_df(data)

#Identify the factors and fix some of the units on variables
data$SampleID <- factor(data$SampleID)
data$Tray <- factor(data$Tray)

mySummary <- ddply(data,.(SampleID),numcolwise(mean))
mySummary <- mutate(mySummary,Force = LoadLbf/0.5)

#Now export this to an Excel sheet so I can manually add the Factor columns
saveFileName <- tclvalue(tkgetSaveFile())
tabName="ForceSummary"
exportToExcel(mySummary,saveFileName,tabName)
