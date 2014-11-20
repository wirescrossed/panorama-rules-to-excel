#!/usr/bin/env python
import requests # Method of getting the XML information from PAN.
import xml.etree.ElementTree as ET # Difficult to use but good XML parser.
import xlsxwriter # Creates an Excel Spreadsheet.
import argparse

class Spreadsheet(object):
    """Create a spreadsheet from the XML document."""
    def __init__(self):
        self.name = None
        self.from_member = ""
        self.to_member = ""
        self.source = ""
        self.destination = ""
        self.application = ""
        self.action = None
        self.description = None
        self.disabled = "no" # Set to no since the PAN might return nothing for permit.
        self.expiration = None

    def writeRowHeaders(self):
        """Write the header row of the spreadsheet."""
        titles = ["Rule Name", "From Zone", "To Zone", "Source", "Destination", "Application", "Action", "Description", "Disabled", "Expiration"]
        i = 0
        for title in titles:
            worksheet.write(0, i, title, bold)
            i += 1

    def setName(self, name):
        """Populate the firewall rule description."""
        self.name = name

    def setFromMember(self, from_member):
        """Set firewall from zone."""
        if not self.from_member == "": # If there are multiple entries add a comma to separate.
            self.from_member += chr(10)
        self.from_member +=str(from_member) # Concatenate each entry.

    def setToMember(self, to_member):
        """Set firewall to zone."""
        if not self.to_member == "": # If there are multiple entries add a comma to separate.
            self.to_member += chr(10)
        self.to_member +=str(to_member) # Concatenate each entry.

    def setSource(self, source):
        """Set firewall from source."""
        if not self.source == "": # If there are multiple entries add a comma to separate.
            self.source += chr(10)
        self.source +=str(source) # Concatenate each entry.

    def setDestination(self, destination):
        """Set firewall to destination."""
        if not self.destination == "": # If there are multiple entries add a comma to separate.
            self.destination += chr(10)
        self.destination +=str(destination) # Concatenate each entry.

    def setApplication(self, application):
        """Set firewall to application."""
        if not self.application == "": # If there are multiple entries add a comma to separate.
            self.application += chr(10)
        self.application +=str(application) # Concatenate each entry.

    def setAction(self, action):
        """Populate the firewall rule action."""
        self.action = action

    def setDescription(self, description):
        """Populate the firewall rule action."""
        self.description = description

    def setDisabled(self, disabled):
        """Populate the firewall rule action."""
        self.disabled = disabled

    def setExpiration(self, expiration):
        """Populate the firewall rule action."""
        self.expiration = expiration

    def writeRow(self, row):
        """Writes row to Excel workbook"""
        # Insert validation later
        worksheet.write(row, 0, self.name, dataformat)
        worksheet.write(row, 1, self.from_member, dataformat)
        worksheet.write(row, 2, self.to_member, dataformat)
        worksheet.write(row, 3, self.source, dataformat)
        worksheet.write(row, 4, self.destination, dataformat)
        worksheet.write(row, 5, self.application, dataformat)
        worksheet.write(row, 6, self.action, dataformat)
        worksheet.write(row, 7, self.description, dataformat)
        worksheet.write(row, 8, self.disabled, dataformat)
        worksheet.write(row, 9, self.expiration, dataformat)

        print "Name: ", self.name
        print "From Zone: ", self.from_member
        print "To Zone: ", self.to_member
        print "Source: ", self.source
        print "Destination: ", self.destination
        print "Application: ", self.application
        print "Action: ", self.action
        print "Disabled: ", self.disabled
        print "Description: ", self.description
        print "Expiration: ", self.expiration
        print "\n"

    def newRow(self):
        """Prepares for new row by clearing variables in class"""
        excelobj.__init__()

def commandlineparser():
    global args
    parser = argparse.ArgumentParser(description='Convert Palo Alto Network Firewall rules from Panorama to Microsoft Excel.', epilog='i.e. pan_to_excel.py --apikey "23j4kl2j34klj2kl4hf5yf" --firewall "Prod firewall 1" --panorama "https://panorama.somewhere.com')
    parser.add_argument('-k', '--apikey', required=True, help='PAN API Token Key')
    parser.add_argument('-f', '--firewall', required=True, help='Firewall Name')
    parser.add_argument('-p', '--panorama', required=True, help='Panorama Managment URL')
    args = parser.parse_args()

if __name__ == '__main__':

    #Get command line arguments
    commandlineparser()

    row = 0 # Used to track which excel row we are on while parsing XML.

    url = "%s/api/?type=config&action=get&xpath=/config/devices/entry[@name=\'localhost.localdomain\']/device-group/entry[@name=\'%s\']/pre-rulebase/security/rules&key=%s" % (args.panorama, args.firewall, args.apikey)

    xml = requests.get(url)

    document = ET.fromstring(xml.content) # Parse the page the firewall returned as a string into the document object.

    workbook = xlsxwriter.Workbook('Firewall_Policies.xlsx') # Create Excel spreadsheet.
    worksheet = workbook.add_worksheet() # Create new worksheet within the spreadsheet.

    bold = workbook.add_format({'bold': True}) # Cell formatting for row header

    dataformat = workbook.add_format() # Cell Formatting for data.
    dataformat.set_align('top')

    excelobj = Spreadsheet()
    excelobj.writeRowHeaders() # Create friendly row headers in the spreadsheet.

    for result in document: # Start after root (result)
        for rules in result: # Start after result (rules)
            for entries in rules: # Start iterating after rules (entries)
                row += 1
                excelobj.setName(name=entries.attrib.get("name")) # Populate the rule description. Used attrib.get since name is a value within the tag.

                for fromzone in entries.findall("from"): # From zone block
                    for members in fromzone.findall("member"): # From zone block - members block
                        excelobj.setFromMember(members.text)

                for tozone in entries.findall("to"): # To zone block
                    for members in tozone.findall("member"): # To zone block - members block
                        excelobj.setToMember(members.text)

                for source in entries.findall("source"): # From source block
                    for members in source.findall("member"): # From source block - members block
                        excelobj.setSource(members.text)

                for destination in entries.findall("destination"): # application block
                    for members in destination.findall("member"): # application block - members block
                        excelobj.setDestination(members.text)

                for application in entries.findall("application"): # application block
                    for members in application.findall("member"): # application block - members block
                        excelobj.setApplication(members.text)

                for action in entries.findall("action"):
                    excelobj.setAction(action.text)

                for description in entries.findall("description"):
                    excelobj.setDescription(description.text)

                for disabled in entries.findall("disabled"):
                    excelobj.setDisabled(disabled.text)

                for expiration in entries.findall("schedule"):
                    excelobj.setExpiration(expiration.text)

                excelobj.writeRow(row) # Write each row to the spreadsheet.
                excelobj.newRow() # Clear old values and start new row.


    workbook.close() # Close the spreadsheet since we are done with it now.

# XML document structure
# <repsonse>
#   <result>
#       <rules>
#           <entry>
#               <from>
#                   <member>from zone</member>
#               </from>
#                <to>
#                   <member>to zone</member>
#                </to>
#               <source>
#                   <member>source network</member>
#               </source>
#               <destination>
#                   <member>destination network</member>
#               </destination>
#               <application>
#                   <member>application</member>
#               </application>
#               <action>
#                   value
#               </action>
#               <description>
#                   value
#               </description>
#               <disabled>
#                   value
#               </disabled>
#           </entry>
#       </rules>
#   </result>
# </repsonse>
