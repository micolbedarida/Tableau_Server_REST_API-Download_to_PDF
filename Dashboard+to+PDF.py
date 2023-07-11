
# coding: utf-8

# In[1]:


# Import necessary Python libraries
import numpy as np
import tableauserverclient as TSC
import pandas as pd
from fpdf import FPDF


# In[2]:


#Read in token name and secret from Excel file
df = pd.read_excel('//Excel_file_path.xlsx')
Token_Name = df['TokenName'][0]
Token_Secret = df['TokenSecret'][0]

filter_values = [0,1,2] #input your filter values here


# In[4]:


#Login Information
server_url = 'https://server_url.online.tableau.com'
site_id = 'site_id'

tableau_auth = TSC.PersonalAccessTokenAuth(Token_Name, Token_Secret, site_id=site_id)
server = TSC.Server(server_url, use_server_version=True)

#ids of the dashboards as published on server
dash_1 = 'fee1be13-f885-481c-832a-d037597074b1' 
dash_2 = 'b4832b12-857f-431c-9494-c3f8d0769a21'
dash_3 = '0e72f646-b430-43cc-aec9-a719b6e3c4f2'


# In[5]:


with server.auth.sign_in(tableau_auth):
    
    cyber_controls = server.views.get_by_id(dash_1)
    image_req_option = TSC.ImageRequestOptions(imageresolution=TSC.ImageRequestOptions.Resolution.High, maxage=1)
    server.views.populate_image(cyber_controls, image_req_option)
    
    with open("Cyber_controls.png", 'wb') as file:
        file.write(cyber_controls.image)
        
    env_overview = server.views.get_by_id(dash_2)
    server.views.populate_image(env_overview, image_req_option)
        
    with open("Env_overview.png", 'wb') as file:
        file.write(env_overview.image)        
        
    metric_drilldown = server.views.get_by_id(dash_3)
    for i in filter_values:  
        image_req_option.vf('CIS Control Number', i)
        server.views.populate_image(metric_drilldown, image_req_option)        
        with open("Metric_drilldown_{0}.png".format(i), 'wb') as file:
            file.write(metric_drilldown.image)


# In[6]:


#find size of dashboard
with open("Cyber_controls.png", "rb") as f:
    data = f.read(24)
    if data.startswith(b'\x89PNG\r\n\x1a\n') and data[12:16] == b'IHDR':
        width = int.from_bytes(data[16:20], byteorder='big')
        height = int.from_bytes(data[20:24], byteorder='big')
    else:
        raise ValueError("Invalid PNG file")


# In[7]:


# create FPDF object
pdf = FPDF(unit="pt", format=[int(width), int(height)])

# set document properties
pdf.set_title("Cyber Controls dashboard")
images = ['Cyber_controls.png', 'Env_overview.png'] +list("Metric_drilldown_{0}.png".format(i) for i in filter_values)


# In[8]:


# loop through each image and add it to the PDF
for i in images:
    image_path = f"{i}"
    pdf.add_page()
    pdf.image(image_path, 0, 0, pdf.w, pdf.h)


# In[9]:


# save the PDF
pdf.output("//output_file_path.pdf", "F")

