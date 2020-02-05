#loading of basic packages
library(readxl)
library(tidyverse)
library(stringr)
library(stringi)

#Reading the excel file
Customerfeedback <- read_excel("C:/Users/gssaruba/Documents/R_Scripts/VOP Analysis/Customerfeedback.xlsx")

#Basic Text Cleaning
Customerfeedback$Location<-as.factor(Customerfeedback$Location)
Customerfeedback<-Customerfeedback[complete.cases(Customerfeedback),]
Customerfeedback$Comments<-trimws(Customerfeedback$Comments,which=c("both"))

Customerfeedback$Comments<-gsub("\n","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("[0-9]","",Customerfeedback$Comments) #not to be removed as few customers have given numbers
Customerfeedback$Comments<-gsub("\r\n\r\n","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub(">","",Customerfeedback$Comments)
#Customerfeedback$Comments<-gsub(")","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("-","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("\r","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("*","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("N/A","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("NA","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("na","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("n/a","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("NIL","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("xx","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("zz","",Customerfeedback$Comments)
Customerfeedback$Comments<-gsub("()","",Customerfeedback$Comments)

#Removing the double quote for character
Customerfeedback$Comments<-sapply(Customerfeedback$Comments, function(x) gsub("\"", "", x))

#Removing blank lines from the dataset
text<-filter(Customerfeedback,!grepl("^\\s*$",Customerfeedback$Comments))

#Encode the Comments text to ASCII from UTF-8
stri_enc_mark(text$Comments)
text$Comments<-stri_encode(text$Comments, "UTF-8", "ASCII")

#Split of Customer feedback - Chennai and Pune
text_chn<-text%>%filter(str_detect(Location,"Chennai"))
text_pun<-text%>%filter(str_detect(Location,"Pune"))


library(syuzhet)
text_chn_sen<-get_sentences(text_chn$Comments)
text_pun_sen<-get_sentences(text_pun$Comments)

head(text_chn_sen,30)

#library(qdap)
#Code to remove stop words 

#text_chn_sen<-rm_stopwords(text_chn_sen,tm::stopwords("english"),separate = FALSE)
#text_chn_sen[1:100]

#Code to correct spelling in the text

# m<-check_spelling_interactive(text_chn_sen)
# preprocessed(m)
# fixit<-attributes(m)$correct
# fixit(text_chn_sen)
# 
# check_text(text_chn_sen)
# View(text_chn_sen)
# 
# head(text_chn_sen,20)
# head(text_pun_sen,20)

library(udpipe)
#model<-udpipe_download_model(language = "english") #Model already downloaded
udmodel_english<-udpipe_load_model("~/R_Scripts/VOP Analysis/english-ewt-ud-2.4-190531.udpipe")

text_chn_df<-as.data.frame(udpipe::udpipe_annotate(udmodel_english,x=text_chn_sen))
text_pun_df<-as.data.frame(udpipe::udpipe_annotate(udmodel_english,x=text_pun_sen))

table(text_chn_df$upos)

library(tidyverse)
#Basic Statistics - Most Occurring Nouns/Adjectives
#Most occurring Nouns
p1<-text_chn_df%>% filter(upos %in% c('NOUN'))%>%
  count(lemma)%>%
  arrange(desc(n))%>%
  head(12)%>%
  ggplot()+geom_bar(aes(reorder(lemma,-n),n),stat="identity")+labs(title="Chennai - Most Commonly occurring Nouns",x="Nouns",y="count of Nouns")+ theme(axis.text.x = element_text(size=10,angle=90, hjust=1))

 
p2<-text_chn_df%>% filter(upos %in% c('ADJ'))%>%
  count(lemma)%>%
  arrange(desc(n))%>%
  head(12)%>%
  ggplot()+geom_bar(aes(reorder(lemma,-n),n),stat="identity")+labs(title="Chennai - Most Commonly occurring Adjectives",x="Adjectives",y="count of Adjectives") + theme(axis.text.x = element_text(size=10,angle=90, hjust=1))

p3<-text_chn_df%>% filter(upos %in% c('VERB'))%>%
  count(lemma)%>%
  arrange(desc(n))%>%
  head(12)%>%
  ggplot()+geom_bar(aes(reorder(lemma,-n),n),stat="identity")+labs(title="Chennai - Most Commonly occurring Verbs",x="Verb",y="count of Verb") + theme(axis.text.x = element_text(size=10,angle=90, hjust=1))

q1<-text_pun_df%>% filter(upos %in% c('NOUN'))%>%
  count(lemma)%>%
  arrange(desc(n))%>%
  head(12)%>%
  ggplot()+geom_bar(aes(reorder(lemma,-n),n),stat="identity")+labs(title="Pune - Most Commonly occurring Nouns",x="Nouns",y="count of Nouns") + theme(axis.text.x = element_text(size=10,angle=90, hjust=1))


q2<-text_pun_df%>% filter(upos %in% c('ADJ'))%>%
  count(lemma)%>%
  arrange(desc(n))%>%
  head(12)%>%
  ggplot()+geom_bar(aes(reorder(lemma,-n),n),stat="identity")+labs(title="Pune - Most Commonly occurring Adjectives",x="Adjectives",y="count of Adjectives") + theme(axis.text.x = element_text(size=10,angle=90, hjust=1))

q3<-text_pun_df%>% filter(upos %in% c('VERB'))%>%
  count(lemma)%>%
  arrange(desc(n))%>%
  head(12)%>%
  ggplot()+geom_bar(aes(reorder(lemma,-n),n),stat="identity")+labs(title="Pune - Most Commonly occurring Verbs",x="Verb",y="count of Verb") + theme(axis.text.x = element_text(size=10,angle=90, hjust=1))


library(cowplot)
plot_grid(p1,p2,p3,q1,q2,q3)

#Identify key words using Rake
stats <- keywords_rake(x = text_chn_df, term = "lemma",group = "doc_id",  
                       relevant = text_chn_df$upos %in% c("NOUN"),ngram_max = 4)
p1<-stats%>%arrange(desc(rake))%>%head(12)%>%ggplot()+geom_bar(aes(reorder(keyword,-rake),rake),stat="identity")+labs(title="Top Keywords identified for GBS Chennai",x="keywords") + theme(axis.text.x = element_text(size=10,angle=90, hjust=1))


stats <- keywords_rake(x = text_pun_df, term = "lemma",group = "doc_id",  
                       relevant = text_pun_df$upos %in% c("NOUN"),ngram_max = 4)
q1<-stats%>%arrange(desc(rake))%>%head(12)%>%ggplot()+geom_bar(aes(reorder(keyword,-rake),rake),stat="identity")+labs(title="Top Keywords identified for GBS Pune",x="keywords")+ theme(axis.text.x = element_text(size=10,angle=90, hjust=1))

plot_grid(p1,q1)

# Find the Co-occurences

cooc1 <- cooccurrence(x = subset(text_chn_df, upos %in% c("NOUN", "ADJ")), 
                     term = "lemma", 
                     group = "doc_id")

cooc2 <- cooccurrence(x = subset(text_pun_df, upos %in% c("NOUN", "ADJ")), 
                      term = "lemma", 
                      group = c("doc_id"))


head(cooc2)

library(igraph)
library(ggraph)
library(ggplot2)
wordnetwork <- head(cooc1, 56)
wordnetwork <- graph_from_data_frame(wordnetwork)
p1<-ggraph(wordnetwork, layout = "fr") +
  geom_edge_link(aes(width = cooc, edge_alpha = cooc), edge_colour = "yellow") +
  geom_node_text(aes(label = name), col = "darkgreen", size = 4) +
  theme_graph(base_family = "Arial Narrow") +
  theme(legend.position = "none") +
  labs(title = "GBS Chennai Cooccurrences within sentence", subtitle = "Nouns & Adjective")

wordnetwork <- head(cooc2, 74)
wordnetwork <- graph_from_data_frame(wordnetwork)
p2<-ggraph(wordnetwork, layout = "fr") +
  geom_edge_link(aes(width = cooc, edge_alpha = cooc), edge_colour = "yellow") +
  geom_node_text(aes(label = name), col = "darkgreen", size = 4) +
  theme_graph(base_family = "Arial Narrow") +
  theme(legend.position = "none") +
  labs(title = "GBS PuneCooccurrences within sentence", subtitle = "Nouns & Adjective")

plot_grid(p1,p2)


#Drawing wordcloud based on unigram, bigram and trigram - GBS Chennai
library(htmltools)
library(wordcloud2)

x<-subset(text_chn_df, upos %in% c("NOUN", "ADJ","VERB","PROPN"))

c1<-data.frame(token=txt_nextgram(x$token,n=1))
y1<-document_term_frequencies(c1,term=c("token"))
y1<-subset(y1,select=c(term,freq))
x1<-y1%>%arrange(desc(freq))%>%head(300)
html_print(wordcloud2(data=x1))


c2<-data.frame(token=txt_nextgram(x$token,n=2))
y2<-document_term_frequencies(c2,term=c("token"))
y2<-subset(y2,select=c(term,freq))
x2<-y2%>%arrange(desc(freq))%>%head(100)
html_print(wordcloud2(data=x2))

c3<-data.frame(token=txt_nextgram(x$token,n=3))
y3<-document_term_frequencies(c3,term=c("token"))
y3<-subset(y3,select=c(term,freq))
x3<-y3%>%arrange(desc(freq))%>%head(25)
html_print(wordcloud2(data=x3))


#Drawing wordcloud based on unigram, bigram and trigram - GBS Pune
x<-subset(text_pun_df, upos %in% c("NOUN", "ADJ","VERB","PROPN"))

c1<-data.frame(token=txt_nextgram(x$token,n=1))
y1<-document_term_frequencies(c1,term=c("token"))
y1<-subset(y1,select=c(term,freq))
x1<-y1%>%arrange(desc(freq))%>%head(300)
html_print(wordcloud2(data=x1))


c2<-data.frame(token=txt_nextgram(x$token,n=2))
y2<-document_term_frequencies(c2,term=c("token"))
y2<-subset(y2,select=c(term,freq))
x2<-y2%>%arrange(desc(freq))%>%head(100)
html_print(wordcloud2(data=x2))

c3<-data.frame(token=txt_nextgram(x$token,n=3))
y3<-document_term_frequencies(c3,term=c("token"))
y3<-subset(y3,select=c(term,freq))
x3<-y3%>%arrange(desc(freq))%>%head(25)
html_print(wordcloud2(data=x3))


#Sentiment Analysis using NRC method 


#Getting sentiment for each sentence using different method - GBS Chennai
syuzhet_vector<-get_sentiment(text_chn_sen,method="syuzhet")
bing_vector<-get_sentiment(text_chn_sen,method="bing")
afinn_vector <- get_sentiment(text_chn_sen, method="afinn")
nrc_vector <- get_sentiment(text_chn_sen, method="nrc", lang = "english")

plot(sign(syuzhet_vector[1:50]),type="l",xlab="Sentence Index",ylab = "Emotional Valence")
plot(sign(bing_vector[1:50]),type="l",xlab="Sentence Index",ylab = "Emotional Valence")
plot(sign(afinn_vector[1:50]),type="l",xlab="Sentence Index",ylab = "Emotional Valence")
plot(sign(nrc_vector[1:50]),type="l",xlab="Sentence Index",ylab = "Emotional Valence")

#Converting it to a dataframe
sentiment_df_Chn<-cbind.data.frame(index=c(1:684),syuzhet_sent=sign(syuzhet_vector),bing_sent=sign(bing_vector),afin_sent=sign(afinn_vector),nrc_sent=sign(nrc_vector))

#Getting sentiment for each sentence using different method - GBS Pune
syuzhet_vector<-get_sentiment(text_pun_sen,method="syuzhet")
bing_vector<-get_sentiment(text_pun_sen,method="bing")
afinn_vector <- get_sentiment(text_pun_sen, method="afinn")
nrc_vector <- get_sentiment(text_pun_sen, method="nrc", lang = "english")

plot(sign(syuzhet_vector[1:50]),type="l",xlab="Sentence Index",ylab = "Emotional Valence")
plot(sign(bing_vector[1:50]),type="l",xlab="Sentence Index",ylab = "Emotional Valence")
plot(sign(afinn_vector[1:50]),type="l",xlab="Sentence Index",ylab = "Emotional Valence")
plot(sign(nrc_vector[1:50]),type="l",xlab="Sentence Index",ylab = "Emotional Valence")

#The get_nrc_sentiment implements Saif Mohammad’s NRC Emotion lexicon
nrc_data_chn<-get_nrc_sentiment(text_chn_sen)
nrc_data_pun<-get_nrc_sentiment(text_pun_sen)

#Plot for eight emotions associated with sentencs as per NRC lexicon
barplot(
  sort(colSums(prop.table(nrc_data_chn[, 1:8]))), 
  horiz = FALSE, 
  cex.names = 0.7, 
  las = 1, 
  main = "Emotions in text - GBS Chennai", xlab="Percentage"
)


barplot(
  sort(colSums(prop.table(nrc_data_pun[, 1:8]))), 
  horiz = FALSE, 
  cex.names = 0.7, 
  las = 1, 
  main = "Emotions in text - GBS Pune", xlab="Percentage"
)

#Plot for positive and negative associated with sentencs as per NRC lexicon

barplot(
  sort(colSums(prop.table(nrc_data_chn[, 9:10]))), 
  horiz = FALSE, 
  cex.names = 0.7, 
  las = 1, 
  main = "Positive & Negative Emotions in  text - GBS Chennai", xlab="Percentage"
)

barplot(
  sort(colSums(prop.table(nrc_data_pun[, 9:10]))), 
  horiz = FALSE, 
  cex.names = 0.7, 
  las = 1, 
  main = "Positive & Negative Emotions in  text - GBS Pune", xlab="Percentage"
)

#Most common words - classification under NRC lexicon
library(tidytext)
library(textdata)
ic<-data.frame(text=text_chn_sen,stringsAsFactors = FALSE)
tidy_ic<-ic%>%unnest_tokens(word,text)
nrc_word_count<-tidy_ic%>%inner_join(get_sentiments("nrc"))%>%count(word,sentiment,sort=TRUE)%>%ungroup()
nrc_word_count %>%
  group_by(sentiment) %>%
  top_n(5) %>%
  ungroup() %>%
  mutate(word = reorder(word, n)) %>%
  ggplot(aes(word, n, fill = sentiment)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~sentiment, scales = "free_y") +
  labs(y = "Contribution to sentiment",
       x = NULL) +
  coord_flip()

#Topic Modelling
library(topicmodels)

