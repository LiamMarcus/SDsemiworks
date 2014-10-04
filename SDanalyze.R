# Load all librarys and load my custom functions
source(file="E:/Projects/Rutility/UtilityFunctions.R")

#Read all the tabs of the excel data file
dataFileName <- tclvalue(tkgetOpenFile())
myData<-importWorksheets(dataFileName)
names(myData)

# select the name of the excel file to save all tables.
saveFileName <- tclvalue(tkgetSaveFile())

#select the worksheet from the object myData
data<-myData$Crush
names(data)

#Identify the factors and fix some of the units on variables
data$SampleID <- factor(data$SampleID)
data$Cavity <- factor(data$Cavity)

# Summarize the data by factors and export the object to a tab the in excel file
mySummary<-ddply(data,.(Cavity),numcolwise(mean))
tabName="CrushSummary"
exportToExcel(mySummary,saveFileName,tabName)

#Plot box plots connected by a line through means
measuredVar="Weight"
summarizeBy=c("Cavity") # separate factors with commas
xlabel="Cavity"
ylabel="grams"
keyName="Cavity"
plotTitle="Part Weight by Cavity"
plotFile=paste(measuredVar, ".boxplot.png",    sep="")
p1<-ggplot(data, aes_string(x=summarizeBy[1], y=measuredVar)) + 
    geom_boxplot() +
    stat_summary(fun.y=mean, geom="line", (aes(group=1))) +
    stat_summary(fun.y=mean, geom="point", size=3, shape=21, fill="white") +
    guides(fill=FALSE) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle)
#    theme_bw()
# send the plot to the Rstudio viewer and make a pdf file
ggsave(p1,width=6,height=4,file=plotFile)
p1

#Plot box plots connected by a line through means
measuredVar="Deflection"
summarizeBy=c("Cavity") # separate factors with commas
xlabel="Cavity"
ylabel="lbf"
keyName="Cavity"
plotTitle="Part Cantilever Deflection by Cavity \n 7-day Aged"
plotFile=paste(measuredVar, ".boxplot.png",    sep="")
p1<-ggplot(data, aes_string(x=summarizeBy[1], y=measuredVar)) + 
    geom_boxplot() +
    stat_summary(fun.y=mean, geom="line", (aes(group=1))) +
    stat_summary(fun.y=mean, geom="point", size=3, shape=21, fill="white") +
    guides(fill=FALSE) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle)
#    theme_bw()
# send the plot to the Rstudio viewer and make a pdf file
ggsave(p1,width=6,height=4,file=plotFile)
p1

#Plot box plots connected by a line through means
measuredVar="TopYield"
summarizeBy=c("Cavity") # separate factors with commas
xlabel="Cavity"
ylabel="lbf"
keyName="Cavity"
plotTitle="Part Top Crush at Yield by Cavity \n 7-day Aged"
plotFile=paste(measuredVar, ".boxplot.png",    sep="")
p1<-ggplot(data, aes_string(x=summarizeBy[1], y=measuredVar)) + 
    geom_boxplot() +
    stat_summary(fun.y=mean, geom="line", (aes(group=1))) +
    stat_summary(fun.y=mean, geom="point", size=3, shape=21, fill="white") +
    guides(fill=FALSE) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle)
#    theme_bw()
# send the plot to the Rstudio viewer and make a pdf file
ggsave(p1,width=6,height=4,file=plotFile)
p1


#select the worksheet from the object myData
data<-myData$Gauge
names(data)

#Identify the factors and fix some of the units on variables
data$Profile <- factor(data$Profile)

tbl_df(data)

#Clean-up of a gauge data table using tidyr
#this call works on a table of gauge data where the locations are column headings
#and the data are the gauge measurements.
data <- data %>%
    gather(location,gauge,R1:C3B4)

#Next, let's change all the gauges to mils using my fixGauge function, but letting
#dplyr's mutate verb apply the function
data <- data %>%
    mutate(gauge =fixGauge(gauge))


gaugeSummary <- summaryCI(data,"gauge",c("Profile","location"))

#Prepare the data for plotting
gaugeSummary$Upper <- gaugeSummary$mean / 2
gaugeSummary$Lower <- 0 - gaugeSummary$Upper
gaugeSummary$ciUpper <- gaugeSummary$ciUpper / 2
gaugeSummary$ciLower <- 0 - gaugeSummary$ciUpper


xlabel="Measurement Location"
ylabel="mils"
plotTitle="Snack Duo Gauge Distribution"
p4<-ggplot(gaugeSummary) +
    geom_ribbon(aes(x=location,ymin=Lower,ymax=Upper, group=1), fill="red", alpha=.4) +
    geom_errorbar(aes(x=location,ymin=ciLower,ymax=ciUpper), colour="black", width=0.1) +
#    facet_grid(SampleID ~ Profile) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme(axis.text.x = element_text(angle = 90))

plotFile=paste("GaugeComparison.CIplot.png",    sep="")
ggsave(p4,width=6,height=4,file=plotFile)
p4
