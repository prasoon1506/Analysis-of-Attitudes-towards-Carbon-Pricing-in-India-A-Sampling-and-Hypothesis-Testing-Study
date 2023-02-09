ds <- read.delim("C:/Users/SOURABH/Downloads/output-onlinecsvtools.txt")


##-----------------------------SECTION A----------------------------------------

#What is the attitude towards carbon pricing?

#I)I)Does gender affect attitude towards carbon pricing?

#creating a subset of all male respondents
df1<-subset(`ds`,gender==2 )

#assigning 1 to agree&strongly_agree and 0 to rest of the responses from male respondents

x=ifelse(`df1`$rich_pay ==1 | `df1`$rich_pay ==5,1,0)
stat.desc(x)

#creating a subset of all female respondents
df2<-subset(`ds`,gender!=2 )

#assigning 1 to agree&strongly_agree and 0 to rest of the responses from female respondents
y=ifelse(`df2`$rich_pay ==1 | `df2`$rich_pay ==5,1,0)
stat.desc(y)

z1 = (mean(y)-mean(x))/sqrt((std.error(x)^2)+(std.error(y))^2)

sd1 = sqrt((std.error(x)^2)+(std.error(y))^2)

#calculating power of proportion test
pwr.t2n.test(sd1,n1 = 80, n2 = 38, sig.level = 0.05, alternative="two.sided") 

#calculating value of n for power=.8
power.t.test(sd1,n = NULL,power=0.8,sig.level=.05,alternative="two.sided")


#II)Does level of education affect attitude towards carbon pricing?

#creating a subset of of all respondents with SS and UG lvl of education
df3<-subset(`ds`,edu ==3 | edu ==2  )

#assigning 1 to agree&strongly_agree and 0 to rest of the responses from respondents
#with  SS and UG lvl of education

a=ifelse(`df3`$rich_pay ==1 | `df3`$rich_pay ==5,1,0)

stat.desc(a)

#creating a subset of of all respondents with PG and PHD lvl of education

df4<-subset(`ds`,edu ==1 | edu ==4  )

#assigning 1 to agree&strongly_agree and 0 to rest of the responses from respondents
#with  PG and PHD lvl of education
b=ifelse(`df4`$rich_pay ==1 | `df4`$rich_pay ==5,1,0)

stat.desc(b)


z2 = (mean(a)-mean(b))/sqrt((std.error(a)^2)+(std.error(b))^2)

sd2 = sqrt((std.error(a)^2)+(std.error(b))^2)

#calculating power of proportion test
pwr.t2n.test(sd2,n1 = 73, n2 = 45, sig.level = 0.05, alternative="two.sided")

#calculating value of n for power=.8
power.t.test(sd2,n = NULL,power=0.8,sig.level=.05,alternative="two.sided")




#III)Does income level affect attitude towards carbon pricing?

#creating a subset of all the respondents with income above 20LPA
df5.1<-subset(`ds`,income ==1)

##assigning 1 to agree&strongly_agree and 0 to rest of the responses from respondents
##income above 20LPA
c = ifelse(`df5.1`$rich_pay ==1 | `df5.1`$rich_pay ==5,1,0)
stat.desc(c)

#creating a subset of all the respondents with income below 5LPA
df5.2<-subset(`ds`,income ==2)

##assigning 1 to agree&strongly_agree and 0 to rest of the responses from respondents
##income above 20LPA
d = ifelse(`df5.2`$rich_pay ==1 | `df5.2`$rich_pay ==5,1,0)
stat.desc(d)

##creating a subset of all the respondents with income b/w 5-20lpa
df5.3<-subset(`ds`,income ==3)

##assigning 1 to agree&strongly_agree and 0 to rest of the responses from respondents
##income b/w 5-20LPA
e=ifelse(`df5.3`$rich_pay ==1 | `df5.3`$rich_pay ==5,1,0)
stat.desc(e)

B = matrix(c(17,18,42,8,16,17),nrow = 2,ncol = 3 ,byrow=T)

rownames(B) = c("agree","disagree")
colnames(B) = c("<5lpa","5-20lpa",">20lpa")
print(B)

Y= chisq.test(B)
print(Y)

w=sqrt(3.2724/(118*2))

#calculating power of chisq test
pwr.chisq.test(w,118,2,.05)

#calculating value of n for power=0.8
pwr.chisq.test(w,N=NULL,2,sig.level=.05,power=0.8)


#IV)Does age affect attitude towards carbon pricing?

#creating a subset of respondents with age b/w 18-24
df7<-subset(`ds`,age ==1)

###assigning 1 to agree&strongly_agree and 0 to rest of the responses from respondents
##with age b/w 18-24
e=ifelse(`df7`$rich_pay ==1 | `df7`$rich_pay ==5,1,0)
stat.desc(e)

#creating a subset of respondents with not b/w 18-24
df8<-subset(`ds`,age !=1)

##assigning 1 to agree&strongly_agree and 0 to rest of the responses from respondents
##with age not b/w 18-24
f=ifelse(`df8`$rich_pay ==1 | `df8`$rich_pay ==5,1,0)
stat.desc(f)

z4= (mean(e)-mean(f))/sqrt((std.error(e)^2)+(std.error(f))^2)

sd4 = sqrt((std.error(e)^2)+(std.error(f))^2)

#calculating power of proportion test
pwr.t2n.test(sd4,n1 = 86, n2 = 32, sig.level = 0.05, alternative="two.sided")

#calculating value of n for power=.8
power.t.test(sd4,n = NULL,power=0.8,sig.level=.05,alternative="two.sided")

#------------------------------------------------------------------------------#


#--------------------------------SECTION B-------------------------------------#


#I)Does gender affect attitude towards carbon tax rate?

#creating a subset of all male respondents
df1<-subset(`ds`,gender==2 )

##assigning 1 to responses that favour less than 5% carbon tax and 0 to responses
##that favour greater than 5% tax rate(male respondents)

g=ifelse(`df1`$amount ==1 | `df1`$amount ==4,1,0)
stat.desc(g)

#creating a subset of all female respondents
df2<-subset(`ds`,gender ==1 )

###assigning 1 to responses that favour less than 5% carbon tax and 0 to responses
##that favour greater than 5% tax rate(female respondents)

h=ifelse(`df2`$amount ==1 | `df2`$amount ==4,1,0)
stat.desc(h)


z5 = (mean(g)-mean(h))/sqrt((std.error(g)^2)+(std.error(h))^2)

sd5 = sqrt((std.error(g)^2)+(std.error(h))^2)

#calculating power of proportion test
pwr.t2n.test(sd5,n1 = 80, n2 = 37, sig.level = 0.05, alternative="two.sided")

#calculating value of n for power=.8
power.t.test(sd5,n = NULL,power=0.8,sig.level=.05,alternative="two.sided")


#II)Does living in city affect the willingness to pay a carbon tax ?

#creating a subset of all respondents who live in village
df2.1<-subset(`ds`,city==2 )

##assigning 1 to responses that favour less than 5% carbon tax and 0 to responses
##that favour greater than 5% tax rate(city respondents)

g.1=ifelse(`df2.1`$amount ==1 | `df2.1`$amount ==4,1,0)
stat.desc(g.1)

#creating a subset of all village respondents
df2.2<-subset(`ds`,city ==1 )

###assigning 1 to responses that favour less than 5% carbon tax and 0 to responses
##that favour greater than 5% tax rate(village respondents)

h.1=ifelse(`df2.2`$amount ==1 | `df2.2`$amount ==4,1,0)
stat.desc(h.1)


z5 = (mean(g.1)-mean(h.1))/sqrt((std.error(g.1)^2)+(std.error(h.1))^2)

sd5 = sqrt((std.error(g.1)^2)+(std.error(h.1))^2)

#calculating power of proportion test
pwr.t2n.test(sd5,n1 = 113, n2 = 4, sig.level = 0.05, alternative="two.sided")

#calculating value of n for power=.8
power.t.test(sd5,n = NULL,power=0.8,sig.level=.05,alternative="two.sided")





#III)Does age affect attitude towards carbon tax rate?

#creating a subset of respondents with age b/w 18-24
df7<-subset(`ds`,age ==1)


##assigning 1 to responses that favour less than 5% carbon tax and 0 to responses
##that favour greater than 5% tax rate(respondents with age b/w 18-24) 

k=ifelse(`df7`$amount ==1 | `df7`$amount ==4,1,0)
stat.desc(k)

#creating a subset of respondents with age not b/w 18-24
df8<-subset(`ds`,age !=1)

##assigning 1 to responses that favour less than 5% carbon tax and 0 to responses
##that favour greater than 5% tax rate(respondents with age not b/w 18-24) 

l=ifelse(`df8`$amount ==1 | `df8`$amount ==4,1,0)
stat.desc(l)

z7= (mean(k)-mean(l))/sqrt((std.error(k)^2)+(std.error(l))^2)

sd7 = sqrt((std.error(k)^2)+(std.error(l))^2)

#calculating power of proportion test
pwr.t2n.test(sd7,n1 = 86, n2 = 32, sig.level = 0.05, alternative="two.sided")

#calculating value of n for power=.8
power.t.test(sd7,n = NULL,power=0.8,sig.level=.05,alternative="two.sided")




#IV)Does income level affect attitude towards carbon tax rate?

#creating a subset of all the respondents with income below 5LPA
df9.1<-subset(`ds`,income ==2)

##assigning 1 to responses that favour less than 5% carbon tax and 0 to responses
##that favour greater than 5% tax rate(respondents with income below 5LPA)
c1 = ifelse(`df9.1`$amount ==1 | `df9.1`$amount ==4,1,0)
stat.desc(c1)

#creating a subset of all the respondents with income b/w 5-20LPA
df9.2<-subset(`ds`,income ==3)

##assigning 1 to responses that favour less than 5% carbon tax and 0 to responses
##that favour greater than 5% tax rate(respondents with income b/w 5-20 LPA)
c2 = ifelse(`df9.2`$amount ==1 | `df9.2`$amount ==4,1,0)
stat.desc(c2)

#creating a subset of all the respondents with income above 20LPA
df9.3<-subset(`ds`,income ==1)

##assigning 1 to responses that favour less than 5% carbon tax and 0 to responses
##that favour greater than 5% tax rate(respondents with income above 20 LPA)
c3 = ifelse(`df9.3`$amount ==1 | `df9.3`$amount ==4,1,0)
stat.desc(c3)

A = matrix(c(31,54,24,3,5,1),nrow = 2,ncol = 3 ,byrow=T)

rownames(A) = c("0-5%","5%+")
colnames(A) = c("<5lpa","5-20lpa",">20lpa")
print(A)

X= chisq.test(A)
print(X)

w2=sqrt(0.611/(118*2))

#calculating power of chisq test
pwr.chisq.test(w2,118,2,.05)

#calculating value of n for power=0.8
pwr.chisq.test(w2,N=NULL,2,sig.level=.05,power=0.8)

#------------------------------------------------------------------------------#


#------------------------------SECTION C---------------------------------------#

#function for calculating Z-stat

z_test <- function(x,mu,n){
  s = sqrt(mu*(1-mu)/n)
  z <- (x-mu)/(s)
  return(z)
}

#Is the sample representative of the Indian Population?


#I)Is the mean age of the sample representative of mean age of Indian population?

#assigning 1 to respondents with age below 24 and 0 to respondents with age above 24
x1=ifelse(`ds`$age ==1 | `ds`$age ==6,1,0)

stat.desc(x1)
#census data on proportion of population b/w the age 18-24

p1 =.4302

sd8 = sqrt(118*p1*(1-p1))

print(sd)

z_test(mean(x1),.4302,118)

#calculating power of Z test
power.t.test(n=118,delta= mean(x1)-p1,sd8,sig.level = .05,type = "one.sample", alternative = "two.sided")

##calculating value of n for power=.8
power.t.test(n=NULL,delta= mean(x1)-p1,sd8,sig.level = .05,power=0.8,type = "one.sample", alternative = "two.sided")



#Is the proportion of gender in the sample representative of Indian population?

##assigning 1 to male respondents and 0 to female respondents
x2=ifelse(`ds`$gender ==2,1,0)

#census data on proportion of male population

p2 = .51511

sd9 = sqrt(117*p2*(1-p2))

z_test(mean(x2),.51511,117)

#calculating power of Z test
power.t.test(n=117,delta= mean(x2)-p2,sd9,sig.level = .05,type = "one.sample", alternative = "two.sided")

##calculating value of n for power=.8
power.t.test(n=NULL,delta= mean(x1)-p1,sd9,sig.level = .05,power=0.8,type = "one.sample", alternative = "two.sided")



#Is the maximum education level in the sample representative of Indian population?

#assigning 1 to respondents with minimum UG level of education 
x3=ifelse(`ds`$edu !=2,1,0)
stat.desc(x3)
#census data on proportion of population with minimum undergraduate level of education

p3 = .1526

sd10 = sqrt(117*p3*(1-p3))

z_test(mean(x3),.1526,118)

#calculating power of Z test
power.t.test(n=118,delta= mean(x3)-p3,sd10,sig.level = .05,type = "one.sample", alternative = "two.sided")

##calculating value of n for power=.8
power.t.test(n=NULL,delta= mean(x1)-p1,sd10,sig.level = .05,power=0.8,type = "one.sample", alternative = "two.sided")



#IV)Is the income level of the respondents representative of Indian population?

#assigning 1 to respondents with income above 20LPA
x4.1=ifelse(`ds`$income ==1,1,0)
stat.desc(x4.1)

#assigning 1 to respondents with income below 5LPA
x4.2=ifelse(`ds`$income ==2,1,0)
stat.desc(x4.2)

#assigning 1 to respondents with income b/w 5-20LPA
x4.3=ifelse(`ds`$income ==3,1,0)
stat.desc(x4.3)

C = matrix(c(sum(x4.1),sum(x4.2),sum(x4.3),87.67,3.07,26.21),nrow = 2,ncol = 3 ,byrow=T)
colnames(C) = c(">20lpa","<5lpa","5-20lpa")
rownames(C) = c("frequency","expected frequency")
print(C)

Z= chisq.test(C)
print(Z)

w3=sqrt(73.28/(118*2))

#calculating power of chisq test
pwr.chisq.test(w3,118,2,.05)

#calculating value of n for power=0.8
pwr.chisq.test(w3,N=NULL,2,sig.level=.05,power=0.8)

