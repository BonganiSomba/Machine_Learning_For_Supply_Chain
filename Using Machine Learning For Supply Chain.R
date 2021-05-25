# Install these packages if they are not already installed.
library(rvest)
library(caret)
library(rpart)
library(ipred)
library(randomForest)
library(kernlab)
library(dplyr)
library(gbm)
library(nnet)
library(mice)
library(lattice)
library(mda)
library(pROC)
library(glmnet)
library(rstudioapi)
library(openxlsx)
library(magrittr)
library(tidyr)
library(RCurl)
library(rockchalk)
library(e1071)
# Sets working directory to be the folder where this .R file is saved
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# Change filename to the one with the actual full data
dat <- openxlsx::read.xlsx("2020 11 09 POM Extract for Elias.xlsx")

names(dat)

# Only keep columns we might use in models
dat <- dat[, c(4:7, 9, 15, 20:22)]

dat$Late.Indicator <- factor(dat$Late.Indicator) %>% relevel(ref = "N")
dat$Merch.Type    <- factor(dat$Merch.Type)
dat$Supplier.Code <- factor(dat$Supplier.Code)
dat$Ship.Date     <- factor(dat$Ship.Date)
dat$PIS.Date      <- factor(dat$PIS.Date)

  
# See how many levels there are...
levels(dat$Supplier.Code)
levels(dat$Late.Indicator)
levels(dat$Supplier.Code)
levels(dat$PIS.Date)
levels(dat$Ship.Date)
levels(dat$Merch.Type)
dat$Transport.Method[dat$Transport.Method == "MANUFACTURER'S OWN"] <- "MANUFACTURERS OWN" 

dat$Transport.Method <- factor(dat$Transport.Method)
dat$Country.of.Origin <- factor(dat$Country.of.Origin)

# Converts 0's to missing values for these two variables
dat$`Ord.(R)`[dat$`Ord.(R)` == 0] <- NA
dat$`Ord.(U)`[dat$`Ord.(U)` == 0] <- NA
dat$`Ord.(R)` <- dat$`Ord.(R)`
dat$`Ord.(U)` <- dat$`Ord.(U)`
##### checking data type####
str(dat)

nrow(dat)

sum(vapply(1:nrow(dat), function(i) any(is.na(dat[i, ])), NA))

# Imputation of missing values
# # This is unnecessary if you don't have missing values in the data.
# See vignette here: https://www.gerkovink.com/miceVignettes/Ad_hoc_and_mice/Ad_hoc_methods.html

# Get information on missing
md.pattern(dat)

# By default, the imputation method used is pmm (predictive mean matching)
# We may need to look at the imputation issue in more detail to get better models;
# this is just a "quick fix" for now
impute <- mice(dat)

# datmice is the data set with all values imputed
datmice <- complete(impute)


# Logistic regression


myformula <- as.formula("Late.Indicator~ Merch.Type + Supplier.Code + 
                        Transport.Method + Country.of.Origin")

# change dat to datmice below if missing data imputation was needed
#####fit <- glm(y ~., data = mydata, family = binomial("logit"), maxit = 100)
myglm <- glm(myformula, 
             data = dat, family = binomial(link = "logit"))

summary(myglm)

### Feature Selection, using LASSO method

# Create matrix with all explanatory variables
x <- dat
x$Late.Indicator <- NULL
x_train <- model.matrix( ~ .-1, x[, c("Merch.Type", "Supplier.Code", 
                                      "Transport.Method", "Country.of.Origin")])

# Runs LASSO method for variable selection using ten-fold cross-validation
# See here: http://web.stanford.edu/~hastie/glmnet/glmnet_alpha.html
# See also: https://eight2late.wordpress.com/2017/07/11/a-gentle-introduction-to-logistic-regression-and-lasso-regularisation-using-r/
set.seed(2)
bestglm <- glmnet::cv.glmnet(x = x_train, y = dat$Late.Indicator, 
                             type.measure = "class", family = "binomial")

# Check graphically how lambda (LASSO tuning parameter) relates 
# to misclassification error (1 - accuracy) 
plot(bestglm)

print(bestglm)
# Accuracy for minimum lambda.min
1 - bestglm$cvm[which.min(bestglm$cvm)]
# Regression coefficient estimates for model with lambda.min 
coef(bestglm, s = "lambda.min")


# Accuracy of lambda.1se
1 - bestglm$cvm[which(bestglm$lambda == bestglm$lambda.1se)]
# Regression coefficient estimates for model with lambda.1se 

coef(bestglm, s = "lambda.1se")

sink(file = "coef.txt")
coef(bestglm, s = "lambda.1se")
names(table(dat$Merch.Type))
names(table(dat$Supplier.Code))
names(table(dat$Transport.Method))
names(table(dat$Country.of.Origin))
closeAllConnections()

dat$Transport.Method <- factor(dat$Transport.Method)
dat$Country.of.Origin <- factor(dat$Country.of.Origin)

dat$Merch.Type <- rockchalk::combineLevels(fac = dat$Merch.Type, 
                                           levs = c("FABIANI", "MARKHAM", "SDNONPLAN NON PLANNING",
                                                    "SPORTSCENE", "TOTALSPORTS"), newLabel = "OTHER")
dat$Merch.Type <- relevel(dat$Merch.Type, ref = "OTHER")
levels(dat$Merch.Type)(dat$Merch.Type)
dat$Supplier.Code <- rockchalk::combineLevels(fac = dat$Supplier.Code, 
                                              levs = c("7", "29", "51", "54", "76", "125","153",
                                                       "207", "247", "263", "298", "1125", "1183","1200",
                                                       "2630", "2714", "2919", "3028", "3313", "3320", "4483",
                                                       "4860", "5083", "5085", "5345", "5429", "5708", "5709",
                                                       "5762", "5780", "5860", "5943", "9001", "10675","11347",
                                                       "14931", "15047", "15811", "16567", "17169", "17265",
                                                       "18241", "18465", "19873", "19922", "20033", "20215", "21096",
                                                       "22155", "22181", "22909", "31096", "33262", "35670", "35671",
                                                       "97170", "200010", "200011", "200014", "200057", "200059", 
                                                       "200064", "200083", "200085", "200094", "200095", "200116", 
                                                       "200121", "200124", "200138", "200141", "200143", "200145", 
                                                       "200156", "200161", "200169", "200170", "200173", "200176",
                                                       "200181", "200182", "200184", "200186", "200187", "200188",
                                                       "200199","200201","200202","200202","200204","200205","200207",
                                                       "200208","200210","200212","200215","218897","300005","300007",
                                                       "300054","300065","300084","300093","300102","300105","300124","300140",
                                                       "300144","300155","300184","300202","300213","300215","300219","300232",
                                                       "300243","300244","300247","300250","300255","300256","300263","300264",
                                                       "300266","300267","300270","300273","300275","300276","300279","300282",
                                                       "300283","300286","300290","300291","300292","300295","300298","300300","300304",
                                                       "300307","300311","300312","300314","300316","300317","300318","300320","300321",
                                                       "300322","300324","300325","330003","330017","330024","330123","330148","330154","330163","330179",
                                                       "330205","330218","330258","330264","330272","330273","330275","330280","330300","330306","330309",
                                                       "330311","330332","330335","330340","330343","330344","330345","330346","330347","330349","330350",
                                                       "330351","330357","330359","330362","330363","330366","330367","330368","330375","330378","330380",
                                                       "330381","330382","330386","330392","330397","330400","330403","330405","330407","330408","330409",
                                                       "330413","330414","330421","330425","330429","330430","330431","330432","330435","330436","330438",
                                                       "330439","330444","330447","330455","330456","330464","330466","330470","330473","330474","330475",
                                                       "330479","330480","330481","330482","330495","330496","330498","330499","330501","330505","330506",
                                                       "400023","400023","400088","400116","400164","400184","400198","400236","400238","400252","400277",
                                                       "400285","400311","400313","400393","400395","400397","400419","400430","400433","400487","400491",
                                                       "400514","400525","400526","400642","400644","400650","400652","400654","400655","400659","400660",
                                                       "400661","400662","400663","400664","400665","400668","400670","400671","401014","402006","402007",
                                                       "402013","402021","402052","402058","402059","402060","402062","402064","402084","402089","402090",
                                                       "402092","402094","402096","402099","402100","402101","402108","402109","402110","402111",
                                                       "402114","402116","402118","402124","402128","402129","402131","402133","402134","402137","402138",
                                                       "402139","402141","402142","402143","420019","420024","420025","420028","420029","420035","420036",
                                                       "420039","420050","420053","420054","420055","420059","420060","420061","420063","420064","420065",
                                                       "420067","440001","440003","440013","440016","440026","440030","440034","440051","440055",
                                                       "440058","440062","440063","440067","440070","440077","440082","440083","440086","440091","440091",
                                                       "440095","440097","440100","440102","440105","440111","500004","500006","500010","500030","500031","500039","500041",
                                                       "500049","500052","500074","500081","500088","500090","500091","500093","500098","500099","500113","500114","500118",
                                                       "500119","500120","500121","500124","500125","500127","500129","500130","500131","500132","600012","600021","600046","600064",
                                                       "600079","600108","600116","600118","600158","600177","600217","600218","600219","600227","600233","600242","600288",
                                                       "600289","600293","600296","600300","600316","600321","600338","600347","600351","600358","600363","600371","600387",
                                                       "600389","600390","600392","600398","600403","600404","600409","600415","600424","600427","600428","600430","600447",
                                                       "600454","600457","600467","600477","600492","600496","600498","600501","600502","600511","600512","600520","600525",
                                                       "600528","600531","600537","600538","600547","600548","600551","600553","600554","600557","600558","600559","600562",
                                                       "600563","600565","600566","600567","600568","600569","600570","600571","600573","600574","600576","600577","600579"), newLabel = "OTHER")

levels(dat$Country.of.Origin)
is.ordered(dat$Country.of.Origin)
str(dat)

dat$Transport.Method <- rockchalk::combineLevels(fac = dat$Transport.Method, 
                                                 levs = c("AIR", "AIR FREIGHT", "MANUFACTURERS OWN",
                                                          "ROAD","SEA", "UNKNOWN"), newLabel = "OTHER OR UNKNOWN")
dat$Transport.Method <- relevel(dat$Transport.Method, ref = "OTHER OR UNKNOWN")

dat$Country.of.Origin <- rockchalk::combineLevels(fac = dat$Country.of.Origin, 
                                                  levs = c("ARGENTINA", "AUSTRIA", "BANGLADESH", "BELGIUM", 
                                                           "BOTSWANA", "BRAZIL", "CAMBODIA", "EL SALVADOR", "ESWATINI", 
                                                           "FINLAND", "GEORGIA", "GERMANY", "HAITI", "HONDURAS", "HONG KONG",
                                                           "INDIA", "INDONESIA", "ISRAEL", "ITALY", "JORDAN", "KAZAKHASTAN", 
                                                           "LESOTHO", "MADAGASCAR", "MALAWI", "MALAYSIA", "MEXICO", "MOROCCO", 
                                                           "MYANMAR", "NAMIBIA", "NICARAGUA", "PAKISTAN", "PHILIPPINES", 
                                                           "POLAND", "SINGAPORE", "SOUTH KOREA", "SRI LANKA", "TAIWAN", 
                                                           "THAILAND", "TUNISIA", "TURKEY", "UK", "UNITED KINGDOM", 
                                                           "UNITED STATES", "USA"), newLabel = "OTHER")
dat$Country.of.Origin <- relevel(dat$Country.of.Origin, ref = "OTHER")

myformula.chosen <- as.formula("Late.Indicator ~ Merch.Type + 
   Country.of.Origin + Transport.Method + Supplier.Code ")

glm.chosen <- glm(myformula.chosen, data = as.data.frame(dat), 
                  family = binomial(link = "logit"))

dat$Supplier.Code <- relevel(dat$Supplier.Code, ref = "OTHER")





# Create a revised formula based on the variables that are included in 
# coef(bestglm, s = "lambda.1se") or coef(bestglm, s = "lambda.min")
# (depending which approach you choose)
# Fill in the names of the variables after the ~ below
myformula.chosen <- as.formula("Late.Indicator ~ Country.of.Origin 
                                + Transport.Method + Supplier.Code ")

glm.chosen <- glm(myformula.chosen, data = as.data.frame(dat), 
               family = binomial(link = "logit"))

# This is the model whose results you should interpret for your output
summary(glm.chosen)


## CROSS VALIDATION.
# Defining the training control

set.seed(2)
fit <- train(myformula.chosen,
             data = dat,
             method = "glm",
             trControl = trainControl(method = "cv", number = 10), 
             family = binomial(link = "logit"))

print(fit)
# This gives you the accuracy (correct prediction %) from your model 
# averaged across the cross-validation procedure
# If it is close to 50%, model is bad; close to 100%, model is good.
fit$results$Accuracy


#############################################################################################
########################                     ##############################################
#######################     cunfuse matrix    ############################################
#####################################################################################
# Fitted probabilities from logistic regression

LateIndicatorglm <- predict.glm(fit, type = "response")

# Fitted responses from logistic regression using threshold of 0.5
# We could try other threshold values
fitLateIndicatorglm <- if_else(LateIndicatorglm >= 0.5, "Y","N")


confusmat.glm <- confusionMatrix(table(dat$Late.Indicator, ), 
                                 positive = "Yes") 
confusmat.glm$table
confusmat.glm$byClass

roc.glm <- roc(datmice$Graduated ~ gradprobglm, plot = TRUE)





# Area under ROC curve should be close to 1 for a good model,
# close to 0.5 for a bad model
roc.glm$auc

### DECISION TREE

# we assume that the same independent variables as were selected 
# for logistic regression should be used here.
output.tree <- rpart::rpart(myformula.chosen, data = dat, 
                            control = rpart.control())

# This will write the decision tree diagram to an image file. This can go in your results, 
# make sure you read up on how to interpret it.
png("decisiontree_lateshipment.png", width = 1280, height = 680)
rpart.plot::rpart.plot(output.tree, box.palette="RdBu", 
                       shadow.col="gray", nn=TRUE)
dev.off()

summary(output.tree)

print(output.tree)

plot(output.tree)
text(output.tree)
set.seed(2)
fit.rpart <- train(myformula.chosen,
                   data = dat,
                   method = "rpart",
                   trControl = trainControl(method = "cv", number = 10))

print(fit.rpart)
fit.rpart$results$Accuracy


### RANDOM FOREST

print(output.RF)
plot(output.RF)

set.seed(2)
fit.rf <- train(myformula.chosen,
                data = dat,
                method = "rf",
                trControl = trainControl(method = "cv", number = 10))

print(fit.rf)
fit.rf$results$Accuracy


# Fitting K - Nearest Neighbor Classifier
# Train the model
myknn <- train(myformula.chosen,
                data = dat, method = "knn", 
               trControl = trainControl(method = "cv", number = 10))

print(myknn)
plot(myknn)

myknn$results$Accuracy


