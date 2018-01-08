from Tkinter import *
from xlwt import *
from xlrd import *
import csv
import random
import os
import datetime
import arcpy
import shutil
import time

class Example(Frame):
    def __init__(self, root):

        Frame.__init__(self, root)
        self.canvas = Canvas(root, borderwidth=0)
        self.frame = Frame(self.canvas)
        self.vsb = Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((4,4), window=self.frame, anchor="nw", 
                                  tags="self.frame")

        self.frame.bind("<Configure>", self.onFrameConfigure)

        self.populate()

    def populate(self):

         
     #Form Entries
         #print("creating widgets")
         self.BlankSpaceLine = Label(self.frame, text="")
         self.BlankSpaceLine.grid(row=0, columnspan=2)

         self.IntroLine = Label(self.frame, text="Welcome: This app will register and create a new GIS project folder for you. Please fill in the information below and click submit. ")
         self.IntroLine.grid(row=1, columnspan=2)
         self.BlankSpaceLine2 = Label(self.frame, text="")
         self.BlankSpaceLine2.grid(row=2)

         self.AnalystNameEntryInstro = Label(self.frame, text="Analyst Name (Typically Your Name)")
         self.AnalystNameEntryInstro.grid(row=3)
         self.AnalystNameEntry = Entry(self.frame)
         self.AnalystNameEntry.grid(row=3, column=1)

         self.AnalystUIDEntrylabel =Label(self.frame, text="Analyst UserID").grid(row=4)
         self.AnalystUIDEntry = Entry(self.frame)
         self.AnalystUIDEntry.grid(row=4, column=1)

         self.ProjNameEntrylabel = Label(self.frame, text="Short Project Name (25 char. max, NO geographic references)").grid(row=5)
         self.ProjNameEntry = Entry(self.frame)
         self.ProjNameEntry.grid(row=5, column=1)

         self.CustomerNameEntrylabel = Label(self.frame, text="Customer Name (POC)").grid(row=6)
         self.CustomerNameEntry = Entry(self.frame)
         self.CustomerNameEntry.grid(row=6, column=1)

         self.CustomerTeamEntrylabel = Label(self.frame, text="Customer Bureau and Team Name (i.e. Global Health/OHA)").grid(row=7)
         self.CustomerTeamEntry = Entry(self.frame)
         self.CustomerTeamEntry.grid(row=7, column=1)

         self.ProjectPurposeEntrylabel = Label(self.frame, text="Project Purpose").grid(row=8)
         self.ProjectPurposeEntry = Text(self.frame, width=35, height=6)
         self.ProjectPurposeEntry.grid(row=8, column=1)

         self.SuccessObjectivesEntrylabel = Label(self.frame, text="Success Measurement Objectives \n (How will this project make a difference?)").grid(row=9)
         self.SuccessObjectivesEntry = Text(self.frame, width=35, height=6)
         self.SuccessObjectivesEntry.grid(row=9, column=1)

         self.Sectorvarlabel= Label(self.frame, text="Sector").grid(row=10)
         self.Sectorvar = StringVar(self.frame)
         self.Sectorvar.set("---Select One---") # initial value
         self.optionList = ("---Select One---", "Ag and Food Security", "Crisis and Conflict","Democracy Human Rights Governance", "Economic Growth and Trade", "Education", "Ending Extreme Poverty", "Environment Climate Change", "Gender Womens Empowerment", "Global Development Lab", "Global Health", "Water and Sanitation")
         self.SectorAbbreviations = ("---Select One---", "AG", "CRISIS","DHRG", "ECON", "ED", "EXTRMPOV", "ENV", "GEND", "LAB", "GH", "WASH")
         self.option = OptionMenu(self.frame, self.Sectorvar, *self.optionList).grid(row=10, column=1)

         def option2_SelectionEvent(event):
			#print(event)
			self.lbox.delete(0, END)
			if event == "Country":
				for x in self.countries:
					self.lbox.insert(END,x[0])
				
			if event == "Regional":
				for x in self.regions:
					self.lbox.insert(END,x[0])
			if event == "Global":
				for x in self.globalentry:
					self.lbox.insert(END,x[0])
					
			
         self.Scalevar = StringVar(self.frame)
         self.Scalevar.set("---Select Geographic Scale---") # initial value
         self.optionList2 = ("---Select Geographic Scale---", "Global", "Regional","Country")
         self.option2 = OptionMenu(self.frame, self.Scalevar, *self.optionList2, command = option2_SelectionEvent).grid(row=11, column=0)

         self.LocationSelectLabel = Label(self.frame, text="\n Select Locations. Multiple selection ARE allowed. \n Shift/Control key not required for multiple selections").grid(row=11, column=1)

         #self.scrollbar = Scrollbar(self)
         #self.scrollbar.pack(side=LEFT, fill=BOTH)

         self.regions = [['Afghanistan and Pakistan', 'Afghanistan and Pakistan'],['Asia', 'Asia'],['Central Asia', 'Central Asia'],['Europe and Eurasia', 'Europe and Eurasia'],['Latin America and the Caribbean', 'Latin America and the Caribbean'],['Middle East', 'Middle East'],['East Africa', 'East Africa'],['West Africa', 'West Africa'],['Southern Africa', 'Southern Africa']]
		 
         self.countries = [['Afghanistan', 'AFG'], ['\xc3\x85land', 'ALA'], ['Albania', 'ALB'], ['Algeria', 'DZA'], ['American Samoa', 'ASM'], ['Andorra', 'AND'], ['Angola', 'AGO'], ['Anguilla', 'AIA'], ['Antarctica', 'ATA'], ['Antigua and Barbuda', 'ATG'], ['Argentina', 'ARG'], ['Armenia', 'ARM'], ['Aruba', 'ABW'], ['Australia', 'AUS'], ['Austria', 'AUT'], ['Azerbaijan', 'AZE'], ['Bahamas', 'BHS'], ['Bahrain', 'BHR'], ['Bangladesh', 'BGD'], ['Barbados', 'BRB'], ['Belarus', 'BLR'], ['Belgium', 'BEL'], ['Belize', 'BLZ'], ['Benin', 'BEN'], ['Bermuda', 'BMU'], ['Bhutan', 'BTN'], ['Bolivia', 'BOL'], ['Bonaire, Saint Eustatius and Saba', 'BES'], ['Bosnia and Herzegovina', 'BIH'], ['Botswana', 'BWA'], ['Bouvet Island', 'BVT'], ['Brazil', 'BRA'], ['British Indian Ocean Territory', 'IOT'], ['British Virgin Islands', 'VGB'], ['Brunei', 'BRN'], ['Bulgaria', 'BGR'], ['Burkina Faso', 'BFA'], ['Burundi', 'BDI'], ['Cambodia', 'KHM'], ['Cameroon', 'CMR'], ['Canada', 'CAN'], ['Cape Verde', 'CPV'], ['Caspian Sea', 'CA-'], ['Cayman Islands', 'CYM'], ['Central African Republic', 'CAF'], ['Chad', 'TCD'], ['Chile', 'CHL'], ['China', 'CHN'], ['Christmas Island', 'CXR'], ['Clipperton Island', 'CL-'], ['Cocos Islands', 'CCK'], ['Colombia', 'COL'], ['Comoros', 'COM'], ['Cook Islands', 'COK'], ['Costa Rica', 'CRI'], ["C\xc3\xb4te d'Ivoire", 'CIV'], ['Croatia', 'HRV'], ['Cuba', 'CUB'], ['Cura\xc3\xa7ao', 'CUW'], ['Cyprus', 'CYP'], ['Czech Republic', 'CZE'], ['Democratic Republic of the Congo', 'COD'], ['Denmark', 'DNK'], ['Djibouti', 'DJI'], ['Dominica', 'DMA'], ['Dominican Republic', 'DOM'], ['East Timor', 'TLS'], ['Ecuador', 'ECU'], ['Egypt', 'EGY'], ['El Salvador', 'SLV'], ['Equatorial Guinea', 'GNQ'], ['Eritrea', 'ERI'], ['Estonia', 'EST'], ['Ethiopia', 'ETH'], ['Falkland Islands', 'FLK'], ['Faroe Islands', 'FRO'], ['Fiji', 'FJI'], ['Finland', 'FIN'], ['France', 'FRA'], ['French Guiana', 'GUF'], ['French Polynesia', 'PYF'], ['French Southern Territories', 'ATF'], ['Gabon', 'GAB'], ['Gambia', 'GMB'], ['Georgia', 'GEO'], ['Germany', 'DEU'], ['Ghana', 'GHA'], ['Gibraltar', 'GIB'], ['Greece', 'GRC'], ['Greenland', 'GRL'], ['Grenada', 'GRD'], ['Guadeloupe', 'GLP'], ['Guam', 'GUM'], ['Guatemala', 'GTM'], ['Guernsey', 'GGY'], ['Guinea', 'GIN'], ['Guinea-Bissau', 'GNB'], ['Guyana', 'GUY'], ['Haiti', 'HTI'], ['Heard Island and McDonald Islands', 'HMD'], ['Honduras', 'HND'], ['Hong Kong', 'HKG'], ['Hungary', 'HUN'], ['Iceland', 'ISL'], ['India', 'IND'], ['Indonesia', 'IDN'], ['Iran', 'IRN'], ['Iraq', 'IRQ'], ['Ireland', 'IRL'], ['Isle of Man', 'IMN'], ['Israel', 'ISR'], ['Italy', 'ITA'], ['Jamaica', 'JAM'], ['Japan', 'JPN'], ['Jersey', 'JEY'], ['Jordan', 'JOR'], ['Kazakhstan', 'KAZ'], ['Kenya', 'KEN'], ['Kiribati', 'KIR'], ['Kosovo', 'KO-'], ['Kuwait', 'KWT'], ['Kyrgyzstan', 'KGZ'], ['Laos', 'LAO'], ['Latvia', 'LVA'], ['Lebanon', 'LBN'], ['Lesotho', 'LSO'], ['Liberia', 'LBR'], ['Libya', 'LBY'], ['Liechtenstein', 'LIE'], ['Lithuania', 'LTU'], ['Luxembourg', 'LUX'], ['Macao', 'MAC'], ['Macedonia', 'MKD'], ['Madagascar', 'MDG'], ['Malawi', 'MWI'], ['Malaysia', 'MYS'], ['Maldives', 'MDV'], ['Mali', 'MLI'], ['Malta', 'MLT'], ['Marshall Islands', 'MHL'], ['Martinique', 'MTQ'], ['Mauritania', 'MRT'], ['Mauritius', 'MUS'], ['Mayotte', 'MYT'], ['Mexico', 'MEX'], ['Micronesia', 'FSM'], ['Moldova', 'MDA'], ['Monaco', 'MCO'], ['Mongolia', 'MNG'], ['Montenegro', 'MNE'], ['Montserrat', 'MSR'], ['Morocco', 'MAR'], ['Mozambique', 'MOZ'], ['Myanmar', 'MMR'], ['Namibia', 'NAM'], ['Nauru', 'NRU'], ['Nepal', 'NPL'], ['Netherlands', 'NLD'], ['New Caledonia', 'NCL'], ['New Zealand', 'NZL'], ['Nicaragua', 'NIC'], ['Niger', 'NER'], ['Nigeria', 'NGA'], ['Niue', 'NIU'], ['Norfolk Island', 'NFK'], ['North Korea', 'PRK'], ['Northern Mariana Islands', 'MNP'], ['Norway', 'NOR'], ['Oman', 'OMN'], ['Pakistan', 'PAK'], ['Palau', 'PLW'], ['Palestina', 'PSE'], ['Panama', 'PAN'], ['Papua New Guinea', 'PNG'], ['Paraguay', 'PRY'], ['Peru', 'PER'], ['Philippines', 'PHL'], ['Pitcairn Islands', 'PCN'], ['Poland', 'POL'], ['Portugal', 'PRT'], ['Puerto Rico', 'PRI'], ['Qatar', 'QAT'], ['Republic of Congo', 'COG'], ['Reunion', 'REU'], ['Romania', 'ROU'], ['Russia', 'RUS'], ['Rwanda', 'RWA'], ['Saint Helena', 'SHN'], ['Saint Kitts and Nevis', 'KNA'], ['Saint Lucia', 'LCA'], ['Saint Pierre and Miquelon', 'SPM'], ['Saint Vincent and the Grenadines', 'VCT'], ['Saint-Barth\xc3\xa9lemy', 'BLM'], ['Saint-Martin', 'MAF'], ['Samoa', 'WSM'], ['San Marino', 'SMR'], ['Sao Tome and Principe', 'STP'], ['Saudi Arabia', 'SAU'], ['Senegal', 'SEN'], ['Serbia', 'SRB'], ['Seychelles', 'SYC'], ['Sierra Leone', 'SLE'], ['Singapore', 'SGP'], ['Sint Maarten', 'SMX'], ['Slovakia', 'SVK'], ['Slovenia', 'SVN'], ['Solomon Islands', 'SLB'], ['Somalia', 'SOM'], ['South Africa', 'ZAF'], ['South Georgia and the South Sandwich Islands', 'SGS'], ['South Korea', 'KOR'], ['South Sudan', 'SSD'], ['Spain', 'ESP'], ['Spratly islands', 'SP-'], ['Sri Lanka', 'LKA'], ['Sudan', 'SDN'], ['Suriname', 'SUR'], ['Svalbard and Jan Mayen', 'SJM'], ['Swaziland', 'SWZ'], ['Sweden', 'SWE'], ['Switzerland', 'CHE'], ['Syria', 'SYR'], ['Taiwan', 'TWN'], ['Tajikistan', 'TJK'], ['Tanzania', 'TZA'], ['Thailand', 'THA'], ['Togo', 'TGO'], ['Tokelau', 'TKL'], ['Tonga', 'TON'], ['Trinidad and Tobago', 'TTO'], ['Tunisia', 'TUN'], ['Turkey', 'TUR'], ['Turkmenistan', 'TKM'], ['Turks and Caicos Islands', 'TCA'], ['Tuvalu', 'TUV'], ['Uganda', 'UGA'], ['Ukraine', 'UKR'], ['United Arab Emirates', 'ARE'], ['United Kingdom', 'GBR'], ['United States', 'USA'], ['United States Minor Outlying Islands', 'UMI'], ['Uruguay', 'URY'], ['Uzbekistan', 'UZB'], ['Vanuatu', 'VUT'], ['Vatican City', 'VAT'], ['Venezuela', 'VEN'], ['Vietnam', 'VNM'], ['Virgin Islands, U.S.', 'VIR'], ['Wallis and Futuna', 'WLF'], ['Western Sahara', 'ESH'], ['Yemen', 'YEM'], ['Zambia', 'ZMB'], ['Zimbabwe', 'ZWE']]

         self.globalentry = [['Global', 'Global']]

         self.ListScrollbar = Scrollbar(self.frame, orient="vertical")
         self.lbox = Listbox(self.frame, width=45, height=25, selectmode='multiple', yscrollcommand=self.ListScrollbar.set)
		 #, yscrollcommand=self.scrollbar.set
         self.ListScrollbar.config(command=self.lbox.yview)
         #self.ListScrollbar.pack(side=RIGHT, fill=Y)
         self.ListScrollbar.grid(row=12, column=2, rowspan=10, sticky=W+E+N+S)
         self.lbox.grid(row=12, column=1, pady=3, rowspan=10)
         #for x in self.countries:
              #self.lbox.insert(END,x[0])   

         #self.scrollbar.config(command=self.lbox.yview)

         self.ResultsHeader_text = StringVar()
         self.ResultsHeader_text.set("Results from System Execution:")
         self.ResultsHeader = Label(self.frame, textvariable=self.ResultsHeader_text).grid(row=12)

         self.Resultslabel_text = StringVar()
         self.Resultslabel_text.set("")
         self.Resultslabel = Label(self.frame, textvariable=self.Resultslabel_text, fg="red", font=("Helvetica", 14)).grid(row=13)

         self.Processinglabel_text = StringVar()
         self.Processinglabel_text.set("")
         self.Processinglabel = Label(self.frame, textvariable=self.Processinglabel_text).grid(row=14)

         self.CVSWriteStatus1_text = StringVar()
         self.CVSWriteStatus1_text.set("")
         self.CVSWriteStatus1 = Label(self.frame, textvariable=self.CVSWriteStatus1_text).grid(row=15)

         self.CVSWriteStatus2_text = StringVar()
         self.CVSWriteStatus2_text.set("")
         self.CVSWriteStatus2 = Label(self.frame, textvariable=self.CVSWriteStatus2_text).grid(row=16)

         self.CVSWriteStatusCombined_text = StringVar()
         self.CVSWriteStatusCombined_text.set("")
         self.CVSWriteStatusCombined = Label(self.frame, textvariable=self.CVSWriteStatusCombined_text).grid(row=17)

         self.WriteDirectoryStatus_text = StringVar()
         self.WriteDirectoryStatus_text.set("")
         self.WriteDirectoryStatus = Label(self.frame, textvariable=self.WriteDirectoryStatus_text).grid(row=18)

         self.WriteDirectoryResult_text = StringVar()
         self.WriteDirectoryResult_text.set("")
         self.WriteDirectoryResult = Label(self.frame, textvariable=self.WriteDirectoryResult_text, font=("Helvetica", 12)).grid(row=19)

         self.WriteMXDStatus_text = StringVar()
         self.WriteMXDStatus_text.set("")
         self.WriteMXDStatus = Label(self.frame, textvariable=self.WriteMXDStatus_text).grid(row=20)

         self.WriteMXDResult_text = StringVar()
         self.WriteMXDResult_text.set("")
         self.WriteMXDResult = Label(self.frame, textvariable=self.WriteMXDResult_text).grid(row=21)

         self.FinalNotes_text = StringVar()
         self.FinalNotes_text.set("")
         self.FinalNotes = Label(self.frame, textvariable=self.FinalNotes_text).grid(row=22)

     #Buttons
         self.button = Button(
              self.frame, text="QUIT", fg="red", command=self.quit
              )
         self.button.grid(row=23, column=0, sticky=W)

         self.create_sheet = Button(self.frame, text="Submit", command=self.check_vals)
         self.create_sheet.grid(row=23, column=1, sticky=W)



    def check_vals(self):
        self.items = map(int, self.lbox.curselection())
        self.Resultslabel_text.set("")
        
        if self.AnalystNameEntry.get() == "":
            self.Resultslabel_text.set("Please enter an analyst name")
        elif self.AnalystUIDEntry.get() == "":
            self.Resultslabel_text.set("Please enter the analyst's ID")
        elif self.ProjNameEntry.get() == "":
            self.Resultslabel_text.set("Please enter a project name")
        elif self.CustomerNameEntry.get() == "":
            self.Resultslabel_text.set("Please enter a customer name")
        elif self.CustomerTeamEntry.get() == "":
            self.Resultslabel_text.set("Please enter the customer's Bureau and Team Name \n (i.e. Global Health/OHA)")
        elif self.ProjectPurposeEntry.get("1.0",'end-1c') == "":
            self.Resultslabel_text.set("Please enter a project purpose")
        elif self.SuccessObjectivesEntry.get("1.0",'end-1c') == "":
            self.Resultslabel_text.set("Please enter a the success objective \n (i.e. how will the project be used to make a difference )")
        elif self.Sectorvar.get() == "---Select One---":
            self.Resultslabel_text.set("Please select a Sector")
        elif self.Scalevar.get() == "---Select Geographic Scale---":
            self.Resultslabel_text.set("Please select a Geographic Scale")
        elif self.Scalevar.get() == "---Select Scale---":
            self.Resultslabel_text.set("Please select a Geographic Scale")
        elif len(self.items) == 0:
            self.Resultslabel_text.set("Please select at least one geographic focus area")
        else:
            self.Resultslabel_text.set("All information provided, processing information...")
            self.process_info()
            
            
        

    def process_info(self):
        self.Processinglabel_text.set("Writing project to spreadsheet.")
        self.items = map(int, self.lbox.curselection())
        self.listselect = self.lbox.curselection()
        self.defquery=""
        self.values = []
        self.GeogItems = []
 
        if self.Scalevar.get() == "Country":
		    for x in self.items:
				self.values.append("'"+self.countries[x][1]+"'")
				self.GeogItems.append(self.countries[x][0])
		    self.defquery= "ISO in ("+", ".join(self.values)+")"
		    self.GeogList="& ".join(self.GeogItems)
        elif self.Scalevar.get() == "Regional":
		    for x in self.items:
				self.values.append("'"+self.regions[x][1]+"'")
				self.GeogItems.append(self.regions[x][0])
		    self.defquery= "USAID_Reg in ("+", ".join(self.values)+")"
		    self.GeogList="& ".join(self.GeogItems)
        elif self.Scalevar.get() =="Global":
		    self.defquery= ""
		    self.GeogList="Global"			
			
		
#writing to csv catalog
        self.now = datetime.datetime.now()
        self.cleanProjName =  re.sub('[^a-zA-Z0-9 \n]', '', self.ProjNameEntry.get())
        if len(self.items) == 1:
            self.ProjFileName = self.lbox.get(self.listselect[0])+'_'+self.cleanProjName+'_'+self.AnalystUIDEntry.get()+'_'+self.now.strftime("%B")+self.now.strftime("%Y")
            self.MXDName = self.lbox.get(self.listselect[0])+'_'+self.cleanProjName
        else:
            self.ProjFileName = self.lbox.get(self.listselect[0])+' et al_'+self.cleanProjName+'_'+self.AnalystUIDEntry.get()+'_'+self.now.strftime("%B")+self.now.strftime("%Y")
            self.MXDName = self.lbox.get(self.listselect[0])+' et al_'+self.cleanProjName


        self.homedir=os.path.dirname(os.path.realpath(__file__))
        self.WriteDirectoryStatus_text.set(self.homedir)
        if self.homedir.endswith('\\'):
            self.homedir=self.homedir
        else:
            self.homedir=self.homedir+'\\'        


        self.writefail = 0
        try:
            with open(self.homedir+'Project_Catalog.csv', 'ab') as (self.csvfile):
                self.catalogcsv = csv.writer(self.csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
                self.listselect = self.lbox.curselection()
                self.catalogcsv.writerow([random.randint(1, 100000), self.AnalystNameEntry.get().replace(",",";"), self.AnalystUIDEntry.get().replace(",",";"), self.ProjNameEntry.get().replace(",",";"), self.CustomerNameEntry.get().replace(",",";"),
                                        self.CustomerTeamEntry.get().replace(",",";"),self.ProjectPurposeEntry.get("1.0",'end-1c').replace(",",";"), self.SuccessObjectivesEntry.get("1.0",'end-1c').replace(",",";"), self.Sectorvar.get(), self.Scalevar.get(),
                                        self.GeogList, self.ProjFileName])
                self.CVSWriteStatus1_text.set("Project successfully written to the project catalog")
        except:
            self.CVSWriteStatus1_text.set("--- NOTE: Failed to add this project to the project catalog. \nPlease notify  of this message---")
            self.writefail = self.writefail + 1

        try:
            with open(self.homedir+'Catalog_Backup.csv', 'ab') as (self.csvfile):
                self.catalogcsv = csv.writer(self.csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
                self.listselect = self.lbox.curselection()
                self.catalogcsv.writerow([random.randint(1, 100000), self.AnalystNameEntry.get(), self.AnalystUIDEntry.get(), self.ProjNameEntry.get(), self.CustomerNameEntry.get(),
                                        self.CustomerTeamEntry.get(),self.ProjectPurposeEntry.get("1.0",'end-1c'), self.SuccessObjectivesEntry.get("1.0",'end-1c'), self.Sectorvar.get(), self.Scalevar.get(),
                                        self.GeogList, self.ProjFileName])
                self.CVSWriteStatus2_text.set("Project successfully written to the backup catalog")
        except:
            self.CVSWriteStatus2_text.set("--- WARNING: Failed to add this project to the backup catalog. ---\n Please capture a screenshot of this program \n and send it.")
            self.writefail = self.writefail + 1
        



     #Creating folder structure
        self.workingdir = os.getcwd()
        self.WriteDirectoryStatus_text.set("Creating New Project Directory")
        

        self.newdir=self.homedir+'Projects\\'+self.ProjFileName
        try:
            os.makedirs(self.newdir)
            self.WriteDirectoryResult_text.set("New Project Directory Created Successfully \n The folder name is: \n"+ self.newdir)
        except Exception, e:
            self.WriteDirectoryResult_text.set("--- ERROR Creating New Project Folder---:\n"+str(e)+"\n If you are uncertain about the reason for this error.")
        os.makedirs(self.newdir+'/Data')
        os.makedirs(self.newdir+'/MXD')
        os.makedirs(self.newdir+'/Outputs')
        os.makedirs(self.newdir+'/Reference')
        os.makedirs(self.newdir+'/Workspace')
        

     #Setup MXD and Data
        self.WriteMXDStatus_text.set("Creating MXD Templates")
        collection = ['portrate','landscape']
        for x in collection:
            try:
                if x == 'portrate':
                    self.mxdname = self.newdir+"/MXD/"+self.MXDName+"_Portrate.mxd"
                    shutil.copy(self.homedir+"Projects/Project_Template_READONLY/MXD/Map_Template_Portrait.mxd", self.mxdname)
                if x == 'landscape':
                    self.mxdname = self.newdir+"/MXD/"+self.MXDName+".mxd"
                    shutil.copy(self.homedir+"Projects/Project_Template_READONLY/MXD/Map_Template.mxd", self.mxdname)
                self.mxd = arcpy.mapping.MapDocument(self.mxdname)
                self.df = arcpy.mapping.ListDataFrames(self.mxd)[0]
                for lyr in arcpy.mapping.ListLayers(self.mxd):
                    if lyr.name == "Countries: Base- GADM":
                        lyr.definitionQuery = self.defquery
                        self.ext = lyr.getExtent()
                        self.df.extent = self.ext
                        lyr.definitionQuery = ''
                    if lyr.name == "Admin 1- GADM" or lyr.name=="Admin 2- GADM" or lyr.name=="Admin 2- GADM" or lyr.name=="Cities and Towns- Natural Earth":
                        if self.Scalevar.get() == "Country":
                            lyr.definitionQuery = self.defquery
                            lyr.visible = True
                    if lyr.name == "Countries: Focus- GADM":
                        if self.Scalevar.get() == "Country" or self.Scalevar.get() == "Regional":
                            lyr.definitionQuery = self.defquery
                            lyr.visible = True
                    if lyr.name == "Country Outlines- GADM":
                        if self.Scalevar.get() == "Country":
                            lyr.visible = True
                self.mxd.save()
                self.create_sheet.destroy()
                self.WriteMXDResult_text.set("MXD Template Created Successfully")
            except Exception, e:            
                self.WriteMXDResult_text.set("--- ERROR Creating MXD Template ---\n"+str(e)+"\n Please capture a screenshot of this program \n and send it .")
            
        self.FinalNotes_text.set("Finished Processing")

    def onFrameConfigure(self, event):
        '''Reset the scroll region to encompass the inner frame'''
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

if __name__ == "__main__":
    root=Tk()
    root.geometry("800x600")
    Example(root).pack(side="top", fill="both", expand=True)
    root.mainloop()
