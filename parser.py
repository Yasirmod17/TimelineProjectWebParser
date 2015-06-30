#########################################################
#  Name: Mohammed Ibrahim (Class of 2017)
#  Project: Python script which requests for user's interest (keyword(s)) and populates       
#             google spreadsheet using the user's email adress, pasword and spreadsheet key
#   Purpose: For intended use in timeline making
########################################################

#Below are the various libraries i imported while making the script. most of them i had to download
# as they dont come with the standard python library
from Tkinter import *
import tkMessageBox
import gdata.spreadsheet.service
import gspread
import feedparser
import urlparse
import urllib
import urllib2
from HTMLParser import HTMLParser
from re import sub
from sys import stderr
from traceback import print_exc
from bs4 import BeautifulSoup
import warnings
global email
global password
global spreadsheet_url
global spreadsheet_key
global spreadsheet_name
global keyword
#At some point, the program warns you about unsafe connection to the external spreadsheet. This is meant to ignre that.
warnings.filterwarnings("ignore")



#This is the GUI
class GUIFramework(Frame):
    
    def __init__(self,master=None):
# Tells the class to Initialize itself
        
       #"""Initialise the base class"""
        Frame.__init__(self,master)
        
      # """Set the Window Title"""
        self.master.title("Request Form")
        
       # """Display the main window"
        #with a little bit of padding"""
        self.grid(padx=10,pady=10)
        self.CreateWidgets()
       
    def CreateWidgets(self):
      #  """Create all the widgets that we need"""
                
       # """Create the Text"""
        self.email = Label(self, text="Enter Email:")
        self.email.grid(row=0, column=0)

        self.Password = Label(self, text="Enter Password:")
        self.Password.grid(row=1, column=0)

        self.SpreadsheetURL = Label(self, text="Enter SpreadsheetURL:")
        self.SpreadsheetURL.grid(row=2, column=0)

        self.SpreadsheetName = Label(self, text="Enter SpreadsheetName:")
        self.SpreadsheetName.grid(row=3, column=0)

        self.Keyword = Label(self, text="Enter Keyword(s):")
        self. Keyword.grid(row=4, column=0)
        
      #  """Create the Entry, set it to be a bit wider"""
        self.email = Entry(self)
        self.email.grid(row=0, column=1, columnspan=3)

        self.password = Entry(self)
        self.password.grid(row=1, column=1, columnspan=3)

        self.spreadsheetURL = Entry(self)
        self.spreadsheetURL.grid(row=2, column=1, columnspan=3)

        self.spreadsheetName = Entry(self)
        self.spreadsheetName.grid(row=3, column=1, columnspan=3)

        self.keyword = Entry(self)
        self.keyword.grid(row=4, column=1, columnspan=3)
        
       # """Create the Button, set the text and the 
      #  command that will be called when the button is clicked"""
        self.btnDisplay = Button(self, text="Submit", command=self.Display)
        self.btnDisplay.grid(row=5, column=4)
        
    def Display(self):
     #  """Called when btnDisplay is clicked, displays the contents of self.enText"""
        rows=[]
        email=self.email.get()
        password=self.password.get()
        words=self.keyword.get()
        keyword=words.replace(" ","+")
        spreadsheet_url= self.spreadsheetURL.get()
        spreadsheet_key=getKey(spreadsheet_url)
        spreadsheet_name= self.spreadsheetName.get()
        tkMessageBox.showinfo("Info", "Your email: %s"  %self.email.get()+"\nYour spreadsheetKey: %s"  %spreadsheet_key+"\nYour SpreadsheetName: %s" %self.spreadsheetName.get()+"\nYour Keyword(s): %s"  %self.keyword.get())
        start(email,password,spreadsheet_name,spreadsheet_key,keyword,rows)



#This Function does the final parsing of the rss feed(s) contained in the feedList array. it rips out the important metadata
#from each entry in the feed and adds that to a dictionary. Each entry's dictionary is in turn added to a master array that holds them all.
def rssFeed(keyword,rows):
    print("Here are the RSS results===>")
    feedList=["http://news.google.com/news?pz=1&cf=all&ned=us&hl=en&q=keyword&output=rss&num=100"]          #A list of rss feeds of interest e.g google news, reddit...
    for i in feedList:
    #Uses the feedparser library to "clean up" the webpage for results creating a list of them.
        myfeed=feedparser.parse(i.replace("keyword",keyword))
        #Loops through the list and makes a dictionary of their data.
        for post in myfeed.entries:                                                                 
            try:
                print(post.title)          #prints out each topic/title it finds
                    #This conditional statement is in place just incase we decide to use the Reddit rss feed (which is a mostly discussion thread) it searches through
##                    #the comment page using the urlFromReddit method for a link to the actual article up for discussion
                if "reddit"in post.link:
                    url=urlFromReddit(post.link)
                    while url==None:
                        url=urlFromReddit(post.link)
                    
                    rows.append({"startdate":dateConvert(post.published),"enddate":"","headline":post.title,"text":htmlContent(url).replace("\n"," ").replace("\t"," "), "media":url})
                else:
                    x=post.link
                    y=x.find("&url=")
                    if y>0:
                        rows.append({"startdate":dateConvert(post.published),"enddate":"","headline":post.title,"text":htmlContent(x[y+5:]).replace("\n"," ").replace("\t"," "), "media":x[y+5:]})
                    else:
                        rows.append({"startdate":dateConvert(post.published),"enddate":"","headline":post.title,"text":htmlContent(x).replace("\n"," ").replace("\t"," "), "media": x})

            except:
                    rows.append({"startdate":dateConvert(post.published),"enddate":"","headline":post.title, "text":"None", "media":post.link})
    #prints a line when its done searching
    print("_"*60)
    return rows

#This method finds the actual article link in the Reddit discussion thread using the beautiful soup library
#it is custom to a reddit rss however as other rss feeds might have a different style


def urlFromReddit(link):
        htmlT=urllib.urlopen(link).read()
        soup=BeautifulSoup(htmlT)
        for tag in soup.find_all(tabindex="1",href=True):
            if tag.text in soup.html.title.text:
                return tag['href']




#This method strips all the HTML tags from a html document effectively returning just content of the "<p>" tags. it is utilized to get the text(body) of the document.
def htmlContent(url):
        string=""
        test=urllib.urlopen(url).read()
        soup=BeautifulSoup(test)
        for tag in soup.findAll("p"):
                string+=(tag.text)

        return string

#This mnethod converts date and time format of the web document into the acceptable format for the timeline(e.g converts "9th june 2014 00:00:00" to
# "07/09/2014 00:00:00"
def dateConvert(date):
    months=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    return (str(months.index(date[8:11])+1)+"/"+ date[5:7]+"/"+date[12:26])


#This method is meant to extract the spreadsheet key from the spreadsheet url. the key is cumbersome to type in so this saves the user time
def getKey(url):
    if "key=" in url:
        y=url.find("key=")
        x=y+4
    else:
        y=url.find("/d/")
        x=y+3
    return url[x:x+44]
        
#This is the Start method where the user is prompted for "keyword(s)", email adress, password and spreadsheet url and then the other methods that search the webpage for
# results are called. This method also populates the spreadsheet with the results found. since google spreadshet API requres its own format of headers, the cells of each column
#are updated then the columns populated. when its done, it calls the fixHeader method which reverts the headers to previous format.

def start(email,password,spreadsheet_name,spreadsheet_key,keyword,rows):
        worksheet_id ='od6'                     #Default value for all spreadsheets
        gc = gspread.login(email, password)
        wks = gc.open(spreadsheet_name).sheet1
        wks.update_acell('A1', "startdate")
        wks.update_acell('C1', "headline")
        wks.update_acell('D1', "text")
        wks.update_acell('E1', "media")    
        newRows=rssFeed(keyword,rows)
        client = gdata.spreadsheet.service.SpreadsheetsService()
        client.debug = True
        client.email = email      
        client.password = password
        client.source = 'some description' 
        client.ProgrammaticLogin()
    
        for row in newRows:
            try:
                client.InsertRow(row, spreadsheet_key, worksheet_id)
            except Exception as e:
                pass
        print("="*90)
        fixHeaders(email,password,spreadsheet_name)



#This methods literally just fixes the headers into the desired timeline maker format

def fixHeaders(email,password,spreadsheet_name):
        gk=gspread.login(email, password)
        wkss = gk.open(spreadsheet_name).sheet1
        wkss.update_acell('A1', "Start Date")
        wkss.update_acell('C1', "Headline")
        wkss.update_acell('D1', "Text")
        wkss.update_acell('E1', "Media")

if __name__ == '__main__':
    guiFrame = GUIFramework()
    guiFrame.mainloop()

