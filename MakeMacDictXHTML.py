# -*- coding: utf-8 -*-
"""
Created on Mon Feb 27 11:21:09 2017
# ---:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::--- #
@author: WayneSun
Perpose: Create custum dictionary for mac.
This version is specific for files download from 國家教育研究院.
You will need to convert the "xls" file to "xlsx" file in advance.
Put this script in same directory and run python 3.

"""
import re
import io
import os
from os import path as oph
import glob
# Libries which is not built-in
import pandas

class MakeDictionaryException(Exception):
    pass

class Make_Dict_XML(object):
    def __init__(self,\
    Dict_Name,\
    csv_directory,\
    plst_DCSDictionaryCopyright="Sun ChengWei (http://lcarbon.idv.tw)",\
    plst_DCSDictionaryManufacturerName="Sun,ChengWei",\
    plst_CFBundleIdentifier=None,\
    plst_CFBundleName=None,\
    css_H1_color="rgb(54,28,33)",\
    css_P_color="rgb(97,71,46)"):
        os.chdir(os.getcwd())
        # Name
        self.dname=str(Dict_Name)
        
        self.csv_dire=csv_directory     
        self.xlsx_file_list=[]
        
        self.workpath=os.getcwd()
        self.xmlName=oph.join(self.workpath,"{}.xml".format(Dict_Name))
        self.cssName=oph.join(self.workpath,"{}.css".format(Dict_Name))
        self.plstName=oph.join(self.workpath,"{}.plst".format(Dict_Name))
        self.make_path=(os.path.join(self.csv_dire,"Makefile"))
        self._f=None #xml output
        self.en=None
        self.zhtw=None
        
        self.plst_DCSDictionaryCopyright=plst_DCSDictionaryCopyright
        self.plst_DCSDictionaryManufacturerName=plst_DCSDictionaryManufacturerName
        self.plst_CFBundleIdentifier=plst_CFBundleIdentifier
        self.plst_CFBundleName=plst_CFBundleName
        
        self.css_H1_color=css_H1_color
        self.css_P_color=css_P_color
        
        if self.plst_CFBundleIdentifier == None:
           self.plst_CFBundleIdentifier="com.CWSun.dictionary.{}".format(self.dname)
        if self.plst_CFBundleName == None:
            self.plst_CFBundleName=self.dname
    # ---:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::--- #        

    def Open_file(self):
        self._f = io.open(self.xmlName,"wt",encoding="utf-8")
    def Close_file(self):
        self._f.close()

    # ---:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::--- #
        
    def Write_header(self):
        self._f.write("""<?xml version="1.0" encoding="UTF-8"?>
        <d:dictionary xmlns="http://www.w3.org/1999/xhtml" xmlns:d="http://www.apple.com/DTDs/DictionaryService-1.0.rng">\n""")
        print("[Info] Creating file : {}".format(self.xmlName))
    def Write_word(self,eng,zhtw):
        self._f.write("""
<d:entry id="{0}" d:title="{0}">
\t<d:index d:value="{0}"/>
\t\t<h1>{0}</h1>
\t\t<p>{1}</p>
</d:entry>
<!-- ===== ===== ===== -->
""".format(eng,zhtw)) 
        self._f.flush()

    def Write_end(self):
        self._f.write("</d:dictionary>")
        print("[Info] File closed : {}".format(self.xmlName))

    # ---:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::--- #
        
    def Write_plist(self):
        with io.open(self.plstName,"wt",encoding="utf-8") as _plst:
            _plst.write("""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>CFBundleDevelopmentRegion</key>
	<string>English</string>
	<key>CFBundleIdentifier</key>
	<string>{0}</string>
	<key>CFBundleName</key>
	<string>{1}</string>
	<key>CFBundleShortVersionString</key>
	<string>1.0</string>
	<key>DCSDictionaryCopyright</key>
	<string>{2}</string>
	<key>DCSDictionaryManufacturerName</key>
	<string>{3}</string>
	<key>DCSDictionaryFrontMatterReferenceID</key>
	<string>front_back_matter</string>
	<key>DCSDictionaryDefaultPrefs</key>
	<dict>
		<key>pronunciation</key>
		<string>0</string>
		<key>display-column</key>
		<string>1</string>
		<key>display-picture</key>
		<string>1</string>
		<key>version</key>
		<string>1</string>
	</dict>
</dict>
</plist>""".format(\
self.plst_CFBundleIdentifier,\
self.plst_CFBundleName,\
self.plst_DCSDictionaryCopyright,\
self.plst_DCSDictionaryManufacturerName))
            _plst.flush()
        print("[OK] File created : {}".format(self.plstName))
            # ---:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::--- #
            
    def Write_Makefile(self):
        
        with io.open(self.make_path,"wt",encoding="utf-8") as _fm:
            _fm.write("""###########################
# You need to edit these values.

DICT_NAME={0} 
DICT_SRC_PATH={1}
CSS_PATH={2}
PLIST_PATH=	{3}

DICT_BUILD_OPTS		=
# Suppress adding supplementary key.
# DICT_BUILD_OPTS		=	-s 0	# Suppress adding supplementary key.

###########################

# The DICT_BUILD_TOOL_DIR value is used also in "build_dict.sh" script.
# You need to set it when you invoke the script directly.

DICT_BUILD_TOOL_DIR	=	"/Developer/Extras/Dictionary Development Kit"
DICT_BUILD_TOOL_BIN	=	"$(DICT_BUILD_TOOL_DIR)/bin"

###########################

DICT_DEV_KIT_OBJ_DIR	=	./objects
export	DICT_DEV_KIT_OBJ_DIR

DESTINATION_FOLDER	=	~/Library/Dictionaries
RM			=	/bin/rm

###########################

all:
	"$(DICT_BUILD_TOOL_BIN)/build_dict.sh" $(DICT_BUILD_OPTS) $(DICT_NAME) $(DICT_SRC_PATH) $(CSS_PATH) $(PLIST_PATH)
	echo "Done."


install:
	echo "Installing into $(DESTINATION_FOLDER)".
	mkdir -p $(DESTINATION_FOLDER)
	ditto --noextattr --norsrc $(DICT_DEV_KIT_OBJ_DIR)/$(DICT_NAME).dictionary  $(DESTINATION_FOLDER)/$(DICT_NAME).dictionary
	touch $(DESTINATION_FOLDER)
	echo "Done."
	echo "To test the new dictionary, try Dictionary.app."

clean:
	$(RM) -rf $(DICT_DEV_KIT_OBJ_DIR)""".format(self.dname,self.xmlName,self.cssName,self.plstName))
            _fm.flush()
            print("[OK] File created : {}".format(self.make_path) )

    # ---:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::--- #
    
    def numSortKey(self,filename):
        ''' Seperate numbers from filename and return a list of numbers for sorting algorithm.
        filename (string) : The name of file you want to split.
        '''
        number = re.compile(r'(\d+)')
        parts = number.split(filename)
        parts[1::2] = map(int,parts[1::2])
        return parts
    
    def EnumerateFlie(self,DirectoryPath,extension="csv"):
        Filelist = glob.glob(os.path.join(DirectoryPath,"*."+extension))
        #print("Glob_path in EnumerateFlie : {}".format(os.path.join(DirectoryPath,"*."+extension)))
        return sorted(Filelist,key=self.numSortKey)
        
    def Read_xlsx(self):
        self.xlsx_file_list=self.EnumerateFlie(self.csv_dire,"xlsx")
        print("[Info] {} xlsx files found.".format(len(self.xlsx_file_list)))
        self.en=None
        self.zhtw=None
        if len(self.xlsx_file_list) == 0:
            raise MakeDictionaryException("[Error] No xlsx file found. The process will terminate.")
        else:
            for i in range(len(self.xlsx_file_list)):
                _xlsx=pandas.read_excel(self.xlsx_file_list[i])
                if i == 0:
                    self.en=_xlsx["英文名稱"]
                    self.zhtw=_xlsx["中文名稱"]
                else:                    
                    self.en=pandas.concat([self.en,_xlsx["英文名稱"]],ignore_index=True)
                    self.zhtw=pandas.concat([self.zhtw,_xlsx["中文名稱"]],ignore_index=True)
        print("[OK] Merged all xlsx files")
        print("[Info] {} records founded.".format(len(self.en)))

    # ---:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::--- #        
    def Make_xml(self):
        print("[Info] Start to create the xml file")
        self.Read_xlsx()
        self.Open_file()
        self.Write_header()
        for i in range(len(self.en)):
            en,zhtw=self.en[i],self.zhtw[i]
            try:
                en=en.replace('"','\\"')
                zhtw=zhtw.replace('"','\\"')
            except Exception as e:
                print("i={}".format(i))
                raise Exception(e)
            self.Write_word(en,zhtw)
        self.Write_end()
        self.Close_file()
        print("[OK] File created : {}".format(self.xmlName))
    # ---:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::--- #
    def Write_css(self):
        with io.open(self.cssName,"wt",encoding="UTF-8") as _css:
            _css.write("""@charset "UTF-8";
@namespace d url(http://www.apple.com/DTDs/DictionaryService-1.0.rng);
d|entry {{
}}
h1	{{
	font-size: 150%;
	color: {0}
}}
p	{{
	font-size: 110%;
	color: {1}
}}
span.column {{
	display: block;
	border: solid 2px #c0c0c0;
	margin-left: 2em;
	margin-right: 2em;
	margin-top: 0.5em;
	margin-bottom: 0.5em;
	padding: 0.5em;
}}
""".format(self.css_H1_color,self.css_P_color))
            _css.flush()
        print("[OK] file created : {}".format(self.cssName))
    
    def Make_all(self):
        self.Make_xml()
        self.Write_css()
        self.Write_plist()
        self.Write_Makefile()
        """        
        self.plst_DCSDictionaryCopyright=plst_DCSDictionaryCopyright
        self.plst_DCSDictionaryManufacturerName=plst_DCSDictionaryManufacturerName
        self.plst_CFBundleIdentifier=plst_CFBundleIdentifier
        self.plst_CFBundleName=plst_CFBundleName
        """

if __name__ == "__main__":
    #===== Check the version of python =====
    import sys
    if sys.version_info[0]<3:
        raise MakeDictionaryException("Must use python 3")
    if sys.platform == 'win32':
        print("\n[Warning] The script is running under Windows, the path of the files will be unable to use on Mac.\n")
    #===== First : customize the information below : ===== 
    Name_of_Dictionary="國家教育研究院地球科學名詞 NAER Geology Dictionary"
    Directory_where_xlsx_is=os.getcwd() # Or something like r"~/Dictionary" or r"D:\\Dictionary"
    Copyright_of_dictionary="Sun ChengWei (http://lcarbon.idv.tw)"
    Manufacturer_Name_of_Dictionary="Sun,ChengWei"
    BundleIdentifier="com.CWSun.dictionary.NAER.Geology"
    FBundleName=None
    Title_Colour="rgb(54,28,33)"
    Content_colour="rgb(97,71,46)"
    #===== Then we will start to create files we need =====
    import time
    print("Pause for 5 sec.\nTo stop the script, press 'ctrl'+ 'c'\n")
    time.sleep(5)
    
    Dic=Make_Dict_XML(Name_of_Dictionary,\
    Directory_where_xlsx_is,\
    Copyright_of_dictionary,\
    Manufacturer_Name_of_Dictionary,\
    BundleIdentifier,\
    FBundleName,\
    Title_Colour,\
    Content_colour)
    Dic.Make_all()
    
    # ===== Finally  =====
    # In terminal, run "make" to compile the program.
    # You will need Xcode for that.
    if sys.platform == 'win32':
        print("\n[Warning] The script is running under Windows, the path of the files will be unable to use on Mac.\n")