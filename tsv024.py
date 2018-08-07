# -*- coding: utf-8 -*-
"""
Created on Thurs Jul 19

Purpose: To automatically populate spreadsheets with pubmed articles

@author: Daniel Alber
twitter: @dalber_
"""

from Bio import Entrez
import openpyxl
import tkinter as tk
from tkinter import filedialog
import os
from time import sleep
import gc

# asks for string input and confirms before proceeding
def confirm_enter_string(prompt):
    not_valid = True
    to_return = ""
    while not_valid:
        try:
            to_return = str(input(prompt))
            check_entry = str(input("Is {} your desired entry? y/n".format(to_return))).lower()
            if check_entry == "y" or check_entry == "yes":
                not_valid = False
        except ValueError:
            clear()
            print("Invalid string, try again")
    return to_return

# file dialog w/o window
def file_dial():
    
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askdirectory()
    return file_path

# clear screen w/ call to 'clear()'
clear = lambda: os.system('cls')

# Article class
class Article():
    id = 0
    email = None
    
    def __init__(self, pmid):
        
        # PMID
        self.pmid = pmid
        
        # errors
        list_of_errors = []
        article_found = True
        
        # initial fetch from entrez/pubmed api + link
        try:
            citation_pull = self.fetch_single_details(pmid)['PubmedArticle'][0]['MedlineCitation']['Article']
            url_temp = 'https://www.ncbi.nlm.nih.gov/pubmed/?term=' + str(self.pmid)
        except IndexError:
            citation_pull = {}
            url_temp = 'URL'
            list_of_errors.append("Article may not have been found ... Review Carefully")
            article_found = False
        self.link = url_temp
        
        # title
        try:
            temp_title = citation_pull['ArticleTitle']
        except:
            temp_title = "TITLE"
            if article_found:
                list_of_errors.append("Title not found")
        self.title = temp_title
        # author
        try:
            author_list_temp = self.all_authors(citation_pull['AuthorList'])
        except KeyError:
            author_list_temp = "AUTHORS"
            if article_found:
                list_of_errors.append("Authors not found")
        self.author = author_list_temp
        
        # date
        try:
            date_int_temp = int(citation_pull['Journal']['JournalIssue']['PubDate']['Year'])
        except ValueError:
            date_int_temp = citation_pull['Journal']['JournalIssue']['PubDate']['Year']
            list_of_errors.append("Non numeric date")
        except KeyError:
            date_int_temp = 0
            if article_found:
                list_of_errors.append("Date not found")
        self.date = date_int_temp
        
        # journal
        try:
            journal_temp = citation_pull['Journal']['Title']
        except KeyError:
            journal_temp = "JOURN"
            if article_found:
                list_of_errors.append("Journal not found")
        self.journal = journal_temp
        
        # language
        try:
            lang_temp = citation_pull['Language'][0]
        except KeyError:
            lang_temp = "LANG"
            if article_found:
                list_of_errors.append("Language not found")
        self.language = lang_temp
        
        # abstract
        try:
            abst_temp = citation_pull['Abstract']['AbstractText'][0]
        except KeyError:
            abst_temp = "ABSTRACT"
            if article_found:
                list_of_errors.append("Abstract not found")
        self.abstract = abst_temp
        
        # error report
        if not list_of_errors:
            self.errors = "No errors"
        else:
            self.errors = ", ".join(list_of_errors)
        
    def all_authors(self, author_list):
        
        # init empty temp list
        author_list_not_joined = []
        
        for author in author_list:
            
            # last name or default LAST
            try:
                last = author['LastName']
            except (KeyError, TypeError):
                last = "LAST"
            
            # initials or default FIRST
            try:
                inits = author['Initials']
            except (KeyError, TypeError):
                inits = "FIRST"
            
            # single author's name LAST INITIALS
            author_string = last + " " + inits
            
            # appends to empty temp list; order is preserved
            author_list_not_joined.append(author_string)
            
        # returns list of authors joined by commas
        return ", ".join(author_list_not_joined)
        
        
    def fetch_single_details(self, pmid):
        Entrez.email = Article.email
        handle = Entrez.efetch(db = 'pubmed',
                              retmode = 'xml',
                              id = pmid)
        results = Entrez.read(handle)
        return results 
    
    def test_article(self):
        print("PMID: {}".format(self.pmid))
        print("Title: {}".format(self.title))
        print("Author(s): {}".format(self.author))
        print("Date: {}".format(self.date))
        print("Journal: {}".format(self.journal))
        print("Language: {}".format(self.language))
        print("Abstract: {}".format(self.abstract))
        print("Link: {}".format(self.link))
        print("Errors: {}".format(self.errors))
        print("\n")
        
# Sheet class
class MakeSheet():
    def __init__(self):
        self.wb = openpyxl.Workbook()
        self.sheet_name = confirm_enter_string("What would you like to name the spreadsheet?")
        self.wb.create_sheet(index = 0, title = self.sheet_name)
        self.sheet = self.wb[self.sheet_name]
        self.init_column_names()
        
    def init_column_names(self):
        self.sheet['A1'] = "PMID"
        self.sheet['B1'] = "Title"
        self.sheet['C1'] = "Authors"
        self.sheet['D1'] = "Year"
        self.sheet['E1'] = "Language"
        self.sheet['F1'] = "Journal"
        self.sheet['G1'] = "Abstract"
        self.sheet['H1'] = "Link"
        
    def insert_article(self, article):
        Article.id += 1
        # PMID
        self.sheet['A{}'.format(Article.id+1)] = article.pmid
        # title
        self.sheet['B{}'.format(Article.id+1)] = article.title
        # authors
        self.sheet['C{}'.format(Article.id+1)] = article.author
        # year
        self.sheet['D{}'.format(Article.id+1)] = article.date
        # language
        self.sheet['E{}'.format(Article.id+1)] = article.language
        # journal
        self.sheet['F{}'.format(Article.id+1)] = article.journal
        # abstract
        self.sheet['G{}'.format(Article.id+1)] = article.abstract
        # link
        self.sheet['H{}'.format(Article.id+1)] = article.link
        
    def save_sheet(self):
        print("\nWhere do you want to save the sheet?")
        directory = file_dial()
        try:
            os.chdir(directory)
            self.wb.save(self.sheet_name + ".xlsx")
            print("\nSheet successfully saved as '{}' in: '{}'".format((self.sheet_name + ".xlsx"), directory))
        except Exception:
            print("\nError ... Sheet not saved")

# Driver class    
class RunProgram():
    # static messages
    welcome_message = "Welcome to Dalber's Pubmed Spreadsheet Generator V 0.24\n\nDeveloped @ the Center for Surgery and Public Health, Brigham and Women's Hospital\n\nPlease report any issues/bugs or suggestions to 'dalber@partners.org'\n\nThank you for using!\n\n"
    help_message = "Commands (more to come in next version!):\nadd - adds studies to sheet\nsave - saves sheet\nhelp - displays this message\nexit - exits program"
    
    def __init__(self):
        clear()
        print(RunProgram.welcome_message)
        Article.email = self.enter_email()
        Article.id = 0
        self.num_articles = 0
        self.make_sheet = MakeSheet()
        self.keep_looping_global = True
        clear()
    
    # enter email - THIS IS CIRCULAR?
    def enter_email(self):
        entered = False
        email = ""
        while not entered:
            email = str(input("\nEnter your email address (in case of pubmed inquiry) ... "))
            y_no = input("\nIs {} your desired entry? (y/n)".format(email))
            clear()
            if y_no == "y" or y_no == "yes":
                entered = True 
        return email
    
    # runs top level    
    def what_to_do(self):
        print(RunProgram.help_message)
        while self.keep_looping_global:
            print("\nSheet name: " + self.make_sheet.sheet_name + "\n")
            command = str(input("\nWhat would you like to do? (type 'help' for list of commands)")).lower()
            clear()
            self.commands(command)
        print("\nTotal articles in sheet: {}\nExiting ...".format(self.num_articles))
        gc.collect()
        sleep(3)            
          
    # command parsing
    def commands(self, cmd):
        if cmd == "exit":
            clear()
            yes_no = str(input("\nIs everything saved? Are you sure you want to exit? (y/n)")).lower()
            if yes_no == "y" or yes_no == "yes":
                self.keep_looping_global = False
        
        elif cmd == "add":
            clear()
            self.add_articles()
            clear()
            
        elif cmd == "save":
            clear()
            self.save_sheet()
            
        elif cmd == "help":
            print(RunProgram.help_message)
    
        else:
            clear()
            print("\nInvalid command, try again (print help for list of commands)")
     
    # adds articles until user types 'done'
    def add_articles(self):
        cont = True
        while cont:
            clear()
            print("\nNumber of articles in sheet: {}".format(self.num_articles))
            pmid = self.init_pmid()
            if pmid == "done":
                cont = False
                clear()
                break
            new_article = Article(pmid)
            print("\n")
            new_article.test_article()
            add_article = str(input("\n\nIs this the article you want to add? (y/n)")).lower()
            if add_article == "y" or add_article == "yes":
                self.num_articles += 1
                self.make_sheet.insert_article(new_article)
    
    # helper function to add pmid/manage input for article adding
    def init_pmid(self):
        valid = False
        error_message = "Invalid PMID, try again"
        pmid = 0
        while not valid:
            # todo - rearrange this try except block
            try:
                pmid = input("\nEnter the next PMID to add or 'done' if finished adding)")
                if pmid == "done":
                    valid = True
                else:
                    pmid = int(pmid)
                    # 8 digit pmid
                    if len(str(pmid)) <= 8:
                        valid = True
                    else:     
                        clear()
                        print(error_message)
            except ValueError:    
                clear()
                print(error_message)
        return pmid
        
    # saves sheet in current state
    def save_sheet(self):
        self.make_sheet.save_sheet()
        
# top level code
if __name__ == "__main__":
    print("\nloading ...\n")
    run_main = RunProgram()
    run_main.what_to_do()

        
        
        
        
