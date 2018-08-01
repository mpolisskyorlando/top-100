# -*- coding: utf-8 -*-
"""
Created on Tue Jul 31 20:24:35 2018

@author: mpolissky
"""


import pandas as pd #for dataframe manipulation
import glob
import os
import shutil
from fuzzywuzzy import fuzz #for fuzzy matching of lines
from xml.etree import ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter.messagebox import showerror

def loadtop100():
    global inputs
    filepath = askopenfilename(title = "Import Top 100 Submits") 
    inputs = pd.read_excel(filepath)
    tk.messagebox.showinfo('Process Complete', 'Process Complete.') 
    return inputs

def loadcontributors():
    global contributors
    filepath = askopenfilename(title = "Import Contributors List") 
    contributors = pd.read_excel(filepath)
    tk.messagebox.showinfo('Process Complete', 'Process Complete.') 
    return contributors


def match_and_save_results():
    #concatenate first and last name from inputs tab
    inputs['name_match'] = inputs.iloc[:,1]+' '+inputs.iloc[:,2]
    inputs['name_match'] = inputs['name_match'].apply(lambda x: x.lstrip().lower())
    contributors['name_match'] = contributors['first_name']+' '+ contributors['last_name']
    contributors['name_match'] = contributors['name_match'].apply(lambda x: x.lstrip().lower())
    
    #iterate over inputs dataframe
    matches = []
    for row in inputs.itertuples():
        scores=[]
        for i in contributors.itertuples():
            scores.append((fuzz.partial_token_sort_ratio(row.name_match, i.name_match),i.name_match))
        match = max(scores,key=lambda item:item[0])
        matches.append(match[1])
    
    #append column with matches to inputs dataframe
    inputs['name_join_key'] = matches
    
    #merge dataframes
    results = inputs.merge(contributors, how='left',left_on =['name_join_key'], right_on = ['name_match'])
    results = results[['Timestamp', 'name_join_key', 'first_name', 'last_name', 'title',
         'bio_slug', 'photo_slug', 'bio_thumbnail', 'bio_alt_thumbnail','bio_body','What is the biggest local or state story this week?','What will be the biggest local/Florida story or issue after this week? ']]
    results.rename(columns={'What is the biggest local or state story this week?':'this_week','What will be the biggest local/Florida story or issue after this week? ':'next_week' },inplace=True)
    
    #remove line breaks and extra spaces
    results.this_week = results.this_week.apply(lambda x: str(x).replace('\n', ' ').replace('\r', '').lstrip())
    results.next_week = results.next_week.apply(lambda x: str(x).replace('\n', ' ').replace('\r', '').lstrip())
    
    #Generate HTML Variables
    looking_ahead = []
    last_week = []
    contributor_title = []
    image_url = []
    full_name = []
    for row in results.itertuples():
        looking_ahead.append(row.next_week)
        last_week.append(row.this_week)
        contributor_title.append(row.title)
        image_url.append(row.bio_thumbnail)
        full_name.append(row.first_name + " " + row.last_name)
    results['looking_ahead']=looking_ahead
    results['last_week']=last_week
    results['contributor_title'] = contributor_title
    results['full_name'] = full_name
    results['image_url'] = image_url
    
    #Create directory to store elementree output
    current_directory = os.getcwd()
    final_directory = os.path.join(current_directory, r'element_tree_output')
    if not os.path.exists(final_directory):
        os.makedirs(final_directory)
    
    # Generate HTML OUTPUT
    i=0 # counter
    for row in results.itertuples():
        i = i + 1
        contributors_title = row.contributor_title
        full_name = row.full_name
        image_url = row.image_url
        last_week = row.last_week
        looking_ahead = row.looking_ahead
    
        #Define all Elements
    
        bio_wrapper = ET.Element('div', attrib={'class': 'fl100-bio'})
        image = ET.Element('img', attrib={'src': image_url, 'class':'fl100-bio-thumbnail'})
    
        text_wrapper = ET.Element('div', attrib={'class': 'fl100-bio-text'})
    
        bioname = ET.Element('p', attrib={'class': 'fl100-bio-name'})
        bioname.text= full_name+', '+contributors_title
    
        par1 = ET.Element('p')
        par2 = ET.Element('p')
    
        biolabel1 = ET.Element('span', attrib={'class': 'fl100-bio-label'})
        biolabel1.text = 'Last week: '
        biolabel2 = ET.Element('span', attrib={'class': 'fl100-bio-label'}) 
        biolabel2.text = 'Looking ahead: '
    
        text1 = ET.Element('span', attrib={'class': 'fl100-bio-text'})
        text1.text = last_week
        text2 = ET.Element('span', attrib={'class': 'fl100-bio-text'})
        text2.text = looking_ahead
    
        #Build out tree
        bio_wrapper.append(image)
        bio_wrapper.append(text_wrapper)
        text_wrapper.append(bioname)
        text_wrapper.append(par1)
        text_wrapper.append(par2)
        par1.append(biolabel1)
        par2.append(biolabel2)
        par1.append(text1)
        par2.append(text2)
    
        ET.ElementTree(bio_wrapper).write(os.path.join(final_directory, r'element'+str(i)+'.txt'), encoding='utf-8',method='html')
    
    
    top_100_list=[r'<style type="text/css">',
                  r'.fl100-bio {border-bottom: 1px solid #666666; padding: 12px 0; margin: 6px 0 24px 0; font-size: 18px; color: #333; line-height: 27px; font-family: Georgia, Times New Roman, serif;}',
                  r'.fl100-bio-thumbnail {float: left; padding: 4px 14px 14px 0; margin: 0; width: 187px; height: 105px;}',
                  r'.fl100-bio-label {font-weight: bold;}',
                  r'@media only screen and (max-width: 600px) {.fl100-bio-thumbnail {float: none;margin: 0 auto;display: block;padding-top: 0;}}',
                  r'</style>'
                 ]
    all_elements = glob.glob(os.path.join(final_directory, "*.txt")) #Get all csv file names
    
    for path in all_elements:
        with open(path,'r', encoding='utf-8') as myfile:
            item = myfile.read()
            top_100_list.append(item)
    
    #remove directory after completed
    shutil.rmtree(final_directory, ignore_errors=True)
    
    savefile = asksaveasfilename(title='Top 100 Results',defaultextension = 'html',initialdir = current_directory)
    
    #Output HTML Results
    with open(savefile, 'w') as f:
        f.write('\n'.join(top_100_list))
    
    tk.messagebox.showinfo('Process Complete', 'Process Complete.') 


root = tk.Tk()
root.title("Florida Top 100")

label_1 = tk.Label(root, text ="Step 1: Load Weekly Submits", font='Helvetica 10 bold', anchor="e")
label_1.grid(row=0, sticky="E")
button_1 = tk.Button(root, text  = "Import", command=loadtop100 )
button_1.grid(row=0, column=1)
label_2 = tk.Label(root, text ="Step 2: Load List of Contributors", font='Helvetica 10 bold', anchor="e")
label_2.grid(row=1, sticky="E")
button_2 = tk.Button(root, text  = "Import", command=loadcontributors )
button_2.grid(row=1, column=1)
label_3 = tk.Label(root, text ="Step 3: Run Match Algorithm and Save Results", font='Helvetica 10 bold', anchor="e")
label_3.grid(row=2, sticky="E")
button_3 = tk.Button(root, text  = "Run & Save Results", command=match_and_save_results )
button_3.grid(row=2, column=1)
root.mainloop()

