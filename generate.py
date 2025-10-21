import pandas as pd
import warnings
from datetime import datetime, timedelta, time
import os
from reportlab.lib.pagesizes import landscape,letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer,PageBreak, Paragraph, Image, PageTemplate, BaseDocTemplate, Frame
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
#from office365.sharepoint.client_context import ClientContext #install the library
#from office365.runtime.auth.user_credential import UserCredential
from io import BytesIO
import requests
from urllib.parse import urlparse

class day_off(object):
    
    def __init__(self, today="") -> None:
       if today=="":
        self.today=datetime.today()
       else: self.today=  datetime.strptime(today, "%Y-%m-%d")

    def easter(self,year):
        if year<1583 or year>2499:
            return None
        tabella={15:(22,2), 16:(22, 2), 17:(23,3), 18:(23, 4), 19:(24,5),
                20:(24,5), 21:(24, 6), 22:(25,0), 23:(26, 1), 24:(25,1)}
        m, n = tabella[year//100]
        a=year%19
        b=year%4
        c=year%7
        d=(19*a+m)%30
        e=(2*b+4*c+6*d+n)%7
        day=d+e
        if (d+e<10):
            day+=22
            month=3
        else:
            day-=9
            month=4
            if ((day==26) or ((day==25) and (d==28) and (e==6) and (a>10))):
                day-=7
        return day, month, year
    
    def working_day(self):
        e=self.easter(self.today.year)
        eas=str(e[0])+'-'+str(e[1])
        easplusone=str(e[0]+1)+'-'+str(e[1])
        holidays=['1-1','6-1','3-2',eas,easplusone,'25-4','1-5','2-6','15-8','1-11','8-12','25-12','26-12']
        weekend=False

        if (str(self.today.day)+'-'+str(self.today.month) in holidays):
            return ("","","","")

        if self.today.weekday()==0: #monday
            yesterday=self.today-timedelta(3)
            tomorrow=self.today+timedelta(1)
        elif self.today.weekday()==4: #friday
            yesterday=self.today-timedelta(1)
            tomorrow=self.today+timedelta(3)
        elif self.today.weekday()==5 or self.today.weekday()==6: weekend=True            
        else:
            yesterday=self.today-timedelta(1)
            tomorrow=self.today+timedelta(1)
        
    
        while(str(tomorrow.day)+'-'+str(tomorrow.month) in holidays):  #this loop checks if tomorrow variable is a holiday. In case, adds 1 day
            tomorrow=tomorrow+timedelta(1)
            if tomorrow.weekday()==5:#saturday
                tomorrow=tomorrow+timedelta(2)
        
        today =str(self.today.year)+"-"+'{:02d}'.format(self.today.month)+"-" +'{:02d}'.format(self.today.day)
        tomorrow =str(tomorrow.year)+"-"+'{:02d}'.format(tomorrow.month)+"-" +'{:02d}'.format(tomorrow.day)
        yesterday =str(yesterday.year)+"-"+'{:02d}'.format(yesterday.month)+"-" +'{:02d}'.format(yesterday.day)
        return (today, tomorrow, yesterday,weekend)


class log_definition(object):
    def __init__(self) -> None:
        pass
    def open_log_file(self,log_file, message):
        if os.path.exists(log_file):
            with (open(log_file,"+a")) as lg:
                lg.write(str(datetime.now()) +" -> "+ message + "\n")
        else: 
            with (open(log_file,"w")) as lg:
                lg.write(str(datetime.now()) + " -> "+ message + "\n")
        return 1
    
    def is_report_uploaded(self,dest_file,today: str) -> bool:
        if os.path.isfile(dest_file+'isloaded.txt'):
            with open(dest_file+'isloaded.txt',"r") as isloaded:
                if str(isloaded.read())==today: #file already loaded
                    return True
                return False
        return False

    def error_message(self,type_error,log_file,lgbk_row=""):
        
        if type_error==1:
            #ctypes.windll.user32.MessageBoxW(0,'Some rows are not completed. Before to create the PDF, please check the Logbook SH','ALERT',16)
            self.open_log_file(log_file+"log.log","Row " +str(lgbk_row)+" field is not completed. Before to create the PDF, please check it.")
        elif type_error==0:
            #ctypes.windll.user32.MessageBoxW(0,'File Logbook SH not found, please check it.','ALERT',16) 
            self.open_log_file(log_file+"log.log","File Logbook SH not found, please check it")
        elif type_error==2:
            #ctypes.windll.user32.MessageBoxW(0,'No training found for today, please check it.','ALERT',16)
            self.open_log_file(log_file+"log.log","No training found for today, please check it.")
        elif type_error==3:
            #ctypes.windll.user32.MessageBoxW(0,'File already loaded.','ALERT',16)  
            self.open_log_file(log_file+"log.log","File already loaded.") 
        elif type_error==4:
            #ctypes.windll.user32.MessageBoxW(0,'File already loaded.','ALERT',16)  
            self.open_log_file(log_file+"log.log",str(lgbk_row))     
        elif type_error==5:
            self.open_log_file(log_file+"log.log","Duration time of row "+lgbk_row+" is not correct") 
        return 1
    
    def analyze_today_rows(self,data,log_file,today):
        for i,row in data[13:].iterrows():
            if str(type(row.iloc[1].date()))!="<class 'datetime.date'>":
                return 1
            if str(row.iloc[1].date())==today:
                #print(str(type(row[1].date())))
                if str(row.iloc[69])=="nan" or row.iloc[69]=="":
                    self.error_message(1,log_file,str(row.iloc[0])+" AJT_ID")
                    return 0
                try:
                    if row.iloc[21] > time(2,0) : #duration
                        self.error_message(5,log_file,str(row.iloc[0]))
                        return 0
                except: pass
                if row.iloc[22] =="" or str(row.iloc[22])=="nan":
                    self.error_message(1,log_file,str(row.iloc[0])+" Outcome")
                    return 0
                if row.iloc[22]!="DCO" and row.iloc[22]!="SDC":
                    if row.iloc[63] =="" or str(row.iloc[63])=="nan":
                        self.error_message(1,log_file,str(row.iloc[0])+" Internal Notes")
                        return 0
                    if row.iloc[23] =="" or str(row.iloc[23])=="nan":
                        self.error_message(1,log_file,str(row.iloc[0])+" Deviation")
                        return 0
        return 1

class data():
    
    def load_file(self, path: str, sheet: str):
        """Load data from an Excel file."""
        try:
            warnings.simplefilter(action='ignore', category=UserWarning)
            return pd.read_excel(path, sheet_name=sheet)
        except Exception as loadErr:
            print(loadErr)
            return pd.array()
    
    def download_from_sharepoint_old(self, site_url, file_url_daily, new_filename, username, password):
        from office365.sharepoint.client_context import ClientContext
        from office365.runtime.auth.user_credential import UserCredential

        # Connessione a SharePoint
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

        try:
            # Scarica il file e lo salva in locale
            with open(new_filename, "wb") as local_file:
                file = ctx.web.get_file_by_server_relative_url(file_url_daily)
                file.download(local_file).execute_query()

            print(f"✅ File scaricato correttamente: {new_filename}")

        except Exception as e:
            print(f"❌ Errore durante il download: {e}") 



    def get_access_token(self,clientID, clientSecret, tenantID):
        """
        Ottiene un token App-Only Azure AD v2 per Microsoft Graph
        """
        token_url = f"https://login.microsoftonline.com/{tenantID}/oauth2/v2.0/token"
        data = {
            "client_id": clientID,
            "client_secret": clientSecret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials"
        }
        r = requests.post(token_url, data=data)
        r.raise_for_status()
        return r.json()["access_token"]

    def get_site_id(self,site_url, access_token):
        """
        Recupera il site_id da Microsoft Graph dato l'URL del sito
        """
        parsed = urlparse(site_url)
        hostname = parsed.netloc
        path = parsed.path.strip("/")  # es. 'sites/GBTS'
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{path}"
        headers = {"Authorization": f"Bearer {access_token}"}
        r = requests.get(graph_url, headers=headers)
        r.raise_for_status()
        return r.json()["id"]

    def download_from_sharepoint__(self,site_url, file_path, new_filename, clientID, clientSecret, tenantID):
        """
        Scarica un file da SharePoint Online via Microsoft Graph API
        """
        try:
            # 1️⃣ Ottieni token
            access_token = self.get_access_token(clientID, clientSecret, tenantID)
            print("✅ Token ottenuto con successo")

            # 2️⃣ Recupera site_id
            site_id = self.get_site_id(site_url, access_token)
            print(f"✅ Site ID ottenuto: {site_id}")

            # 3️⃣ Scarica il file
            graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/content"
            headers = {"Authorization": f"Bearer {access_token}"}
            site_url__old = "https://graph.microsoft.com/v1.0/sites/lcajt.sharepoint.com:/sites/GBTS"
            
            r = requests.get(graph_url, headers=headers, stream=True)
            print(r.status_code)
            print(r.json())
            
            r.raise_for_status()

            with open(new_filename, "wb") as f:
                f.write(r.content)
            print(r.content[:200])
            print(f"✅ File scaricato correttamente: {new_filename}")

        except requests.HTTPError as e:
            print(f"❌ HTTP Error: {e.response.status_code} {e.response.text}")
        except Exception as e:
            print(f"❌ Errore: {e}")

    def download_from_sharepoint(self, site_url, file_path, new_filename, clientID, clientSecret, tenantID):
        """
        Scarica un file da SharePoint Online via Microsoft Graph API. clientSecret scade il 20/10/2027
        """
        try:
            # 1️⃣ Ottieni token
            access_token = self.get_access_token(clientID, clientSecret, tenantID)
            print("Token successfully obtained")

            # 2️⃣ Recupera site_id
            site_id = self.get_site_id(site_url, access_token)
            #print(f"Site ID obtained: {site_id}")

            # 3️⃣ Recupera tutti i drive del sito
            drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
            headers = {"Authorization": f"Bearer {access_token}"}
            drives_response = requests.get(drives_url, headers=headers)
            drives_response.raise_for_status()
            drives = drives_response.json()["value"]

            # 4️⃣ Trova il drive "Documents"
            documents_drive = next((d for d in drives if d["name"] == "Documents"), None)
            if not documents_drive:
                raise Exception("Drive 'Documents' not found in the site.")

            drive_id = documents_drive["id"]
            #print(f"Drive ID trovato: {drive_id}")

            # 5️⃣ URL di download del file
            graph_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path}:/content"
            #print(f"📥 Download da: {graph_url}")

            r = requests.get(graph_url, headers=headers, stream=True)
            #print(f"➡️ Status code: {r.status_code}")

            if r.status_code == 200:
                with open(new_filename, "wb") as f:
                    f.write(r.content)
                print(f"File downloaded: {new_filename}")
            else:
                print(f"Error {r.status_code}: {r.text[:300]}")

        except requests.HTTPError as e:
            print(f"HTTP Error: {e.response.status_code} {e.response.text}")
        except Exception as e:
            print(f"Error: {e}")



        
    def insert_newline_every_n_spaces(self,text, n=3):
        words = text.split()
        new_line = ''
        count = 0
        for word in words:
            new_line += word + ' '
            count += 1
            if count == 3:
                new_line += '\n'
                count = 0
        return new_line.strip() 
    
    def hour_minute_converter(self,total_time:datetime):
        total_hours=total_time.days*24 + int(total_time.seconds/3600)
        total_minutes = int((total_time.seconds%3600)/60)
        return "{:02d}:{:02d}".format(total_hours,total_minutes)

    def read_today_data(self,lgbk,today):
        
        today_sessions=[]
        scheduled_sessions=0
        executed_sessions=0
        outcome_colors={'RSLD':'gray','CANC':'red','SMC':'red','DNCO':'orange','SDNC':'orange','DCO':'white','SDC':'white','ERR':'white'}
        pld_duration=act_duration=timedelta(hours=0,minutes=0,seconds=0)
        for i,row in lgbk[12:].iterrows():
            
            if len(str(row.iloc[1].date()).split('-'))!=3:
                break
            if str(row.iloc[1].date())==today:
                try:
                    
                    sim=str(row.iloc[10])
                    if "LVC" in str(row.iloc[15]):
                        sim=str(row.iloc[10]) + " LVC"
                    
                    ip_notes= "<b>IP Notes:</b> " + self.insert_newline_every_n_spaces(str(row.iloc[24]))
                    internal_notes="<br/><b>Tech Notes:</b> " + self.insert_newline_every_n_spaces(str(row.iloc[63]))
                    if str(row.iloc[24]) == "nan":
                        ip_notes=""
                    if str(row.iloc[63]) == "nan":
                        internal_notes=""
                    styles = getSampleStyleSheet()
                    custom_style = ParagraphStyle(
                                                    'CustomStyle',
                                                    parent=styles['Normal'],  
                                                    fontSize=5,              
                                                    )
                   
                    formatted_text = Paragraph(ip_notes + internal_notes, custom_style)
                    
                    tmp=[row.iloc[69],row.iloc[15],(str(row.iloc[4])+"@\n"+str(row.iloc[3])).replace("nan@nan",""),
                         str(row.iloc[6]).replace("nan",""),sim,row.iloc[11],str(row.iloc[21]).replace("nan",""),row.iloc[22],str(row.iloc[23]).replace("nan","N/A"),\
                         formatted_text,\
                         row.iloc[18],outcome_colors[row.iloc[22]]]
                    scheduled_sessions+=1
                    if str(row.iloc[22]) == "DCO" or str(row.iloc[22]) == "SDC":
                        executed_sessions+=1
                    today_sessions.append(tmp)
                    try:
                        pld_duration=pld_duration+timedelta(hours=row.iloc[18].hour,minutes=row.iloc[18].minute,seconds=row.iloc[18].second)
                        act_duration=act_duration+timedelta(hours=row.iloc[21].hour,minutes=row.iloc[21].minute, seconds=row.iloc[21].second)
                    except: pass #case of duration not available
                except Exception as readingErr: 
                    print(readingErr)
                    return 
        return today_sessions,self.hour_minute_converter(pld_duration),self.hour_minute_converter(act_duration),scheduled_sessions,executed_sessions        


    def read_tomorrow_data(self,lgbk,tomorrow):
        tomorrow_sessions=[]
        pld_duration=timedelta(hours=0,minutes=0,seconds=0)
        for i,row in lgbk[10:].iterrows():
            if str(row.iloc[1].date())==tomorrow:
                sim=str(row.iloc[10])
                if "LVC" in str(row.iloc[15]):
                    sim=str(row.iloc[10]) + " LVC"
                tmp=[row.iloc[69],row.iloc[15],str(row.iloc[4])+"@\n"+str(row.iloc[3]),row.iloc[6],sim,row.iloc[11],row.iloc[18],str(row.iloc[63]).replace("nan","N/A")]
                tomorrow_sessions.append(tmp)
                try:
                    pld_duration=pld_duration+timedelta(hours=row.iloc[18].hour,minutes=row.iloc[18].minute,seconds=row.iloc[18].second)
                except: pass
        return tomorrow_sessions, self.hour_minute_converter(pld_duration)
    
    def read_rtms_data(self,rtms,today):
        rtms_session=[]
        for i,row in rtms[140:].iterrows():
            try:
                if str(row.iloc[1].date())==today:
                    rtms_session.append([row.iloc[0],row.iloc[3],
                                         str(row.iloc[7]).replace("nan",""),
                                         str(row.iloc[8]).replace("nan",""),
                                         row.iloc[10],
                                         self.insert_newline_every_n_spaces(str(row.iloc[12]).replace("nan","")),
                                         self.insert_newline_every_n_spaces(str(row.iloc[14]).replace("nan","")),
                                         self.insert_newline_every_n_spaces(str(row.iloc[15]).replace("nan",""))
                                        ])
            except: pass
        return rtms_session
    

    def read_sim_status(self,lgbk):
        sim_status=[]
        status_color={'OK':'green','NOK':'red','POK':'orange'}
        for i,row in lgbk[0:7].iterrows():
            sim_status.append([row.iloc[0],row.iloc[1],
                               #str(row.iloc[2]).replace("nan","").replace("NaT", ""),
                               "" if pd.isna(row.iloc[2]) else str(row.iloc[2]),
                               str(row.iloc[3]).replace("nan","").replace("NaT", ""),
                               str(row.iloc[4]).replace("nan","").replace("NaT", ""),
                               str(row.iloc[5]).replace("nan","").replace("NaT", "") ,
                               status_color[row.iloc[1]]]) 
        return sim_status
    
    def read_cbt_sbt_status(self,cbt_sbt):
        sbt=[]
        cbt=[]
        status_color={'OK':'green','NOK':'red','POK':'orange'}
        check_sbt=check_cbt=False
        for i,row in cbt_sbt[5:].iterrows():
            if check_sbt and check_cbt:
                sbt.append(["SBTs Availability",str(cbt_sbt.iloc[2,8])+" OK","","","","","green"])
                cbt.append(["CBTs Availability",str(cbt_sbt.iloc[2,18])+" OK","","","","","green"])
                return sbt,cbt
            if not check_sbt and row.iloc[7] != "Totale complessivo":
                sbt.append([row.iloc[7],self.insert_newline_every_n_spaces(row.iloc[8]),
                            self.insert_newline_every_n_spaces(row.iloc[9]),
                            str(row.iloc[10]),"","",status_color[str(row.iloc[8])]])
            else: check_sbt=True
            if not check_cbt and row.iloc[17] != "Totale complessivo":
                cbt.append([row.iloc[17],self.insert_newline_every_n_spaces(row.iloc[18]),
                            self.insert_newline_every_n_spaces(row.iloc[19]),
                           str(row.iloc[20]),"","",status_color[str(row.iloc[18])]])
            else: check_cbt=True

    def  read_cbt_sbt_status_new(self,cbt_sbt, today):
        sbt=[]
        cbt=[]
        sbt_status=18
        cbt_status=32
        status_color={'OK':'green','NOK':'red','POK':'orange'}
        
        try:
            for i,row in cbt_sbt[5:].iterrows():      
                    """if time.strptime(str(row.iloc[1].date()),"%Y-%m-%d")>time.strptime(today,"%Y-%m-%d"):
                        break"""
                    
                    if (row.iloc[6]!= "SBT Room1" and row.iloc[6]!= "SBT Room2") and "SBT" in row.iloc[6] and row.iloc[3]!="OK":
                        sbt_status-=1
                        sbt.append([row.iloc[6],row.iloc[3],row.iloc[8],row.iloc[2],"","",status_color[str(row.iloc[3])]])  

                    elif row.iloc[6] == "SBT Room1" and row.iloc[3]!="OK":
                        sbt_status-=8
                        sbt.append([row.iloc[6],row.iloc[3],row.iloc[8],row.iloc[2],"","",status_color[str(row.iloc[3])]])

                    elif row.iloc[6] == "SBT Room2" and row.iloc[3]!="OK":
                        sbt_status-=10
                        sbt.append([row.iloc[6],row.iloc[3],row.iloc[8],row.iloc[2],"","",status_color[str(row.iloc[3])]])

                    if (row.iloc[6]!= "CBT Room1" and row.iloc[6]!= "CBT Room2") and "CBT" in row.iloc[6] and row.iloc[3]!="OK":
                        cbt_status-=1
                        cbt.append([row.ilocow[6],row.iloc[3],row.iloc[9],row.iloc[2 ],"","",status_color[str(row.iloc[3])]])
                    elif (row.iloc[6] == "CBT Room1" or row.iloc[6] == "CBT Room2") and row.iloc[3]!="OK":
                        cbt_status-=16
                        sbt.append([row.iloc[6],row.iloc[3],row.iloc[8],row.iloc[2],"","",status_color[str(row.iloc[3])]])

                    for sub in sbt:
                        if sub[0] == row.iloc[6] and row.iloc[3]=="OK":
                            sbt.remove(sub)
                            if sub[0] =="SBT Room1":
                                sbt_status+=8
                            elif sub[0] =="SBT Room2":
                                sbt_status+=10
                            else: sbt_status+=1
                    for sub in cbt:
                        if sub[0] == row.iloc[6] and row.iloc[3]=="OK":
                            cbt.remove(sub)
                            if sub[0] =="CBT Room1" or sub[0] =="CBT Room2":
                                sbt_status+=16
                            else: cbt_status+=1
           
        except: pass

        sbt.append(["SBTs Availability",str(sbt_status)+"/18 OK","","","","","green"])
        cbt.append(["CBTs Availability",str(cbt_status)+"/32 OK","","","","","green"])
        return sbt,cbt  

    def read_mpds_status(self,mpds_items):
        mpds=[]
        #status_color={'OK':'green','NOK':'red','POK':'orange'}
        ok=18
        for i,row in mpds_items[1:].items():
            #print(str(i)+": "+str(row[1])+" - "+str(row[24]))
            if "DECIMO" in str(row.iloc[0]).upper():
                if "INOP" in str(row.iloc[0]):
                    mpds.append([str(i),str(row.iloc[24]),"","","","red"])
                    ok-=1
            elif str(i).startswith("MPDS"):
                mpds.append([str(i),"Deployed in "+str(row.iloc[0]),"","","","gray"])
                ok-=1
        mpds.append(["MPDS Station Availability",str(ok)+"/18 OK","","","","","green"])   
        return mpds


class pdf_dev:

    def table_generator_old(self,doc,data,style,percentage=0.10,n_columns=10,n_different_width=1,table_type=1,rowHeights=12): 
        colWidths=[]
        for i in range(0,n_columns):
            if i>=n_columns-n_different_width:
                colWidths.append(doc.width*(1-((n_columns-1)*percentage)))
            else: colWidths.append(doc.width*percentage)
        
        if table_type==1: #table without the cell's height not set
           table=Table(data, colWidths=colWidths)
        else: table=Table(data, colWidths=colWidths,rowHeights=rowHeights)
        table.setStyle(TableStyle(style))
        return  table
    
    def table_generator(self, doc, data, style, percentage=0.10, n_columns=10, n_different_width=1, table_type=1, rowHeights=12):
        colWidths = []
        
        if table_type == 1:
            # Distribuzione standard
            for i in range(0, n_columns):
                if i >= n_columns - n_different_width:
                    colWidths.append(doc.width * (1 - ((n_columns - 1) * percentage)))
                else:
                    colWidths.append(doc.width * percentage)
        else:
            # Riduci la larghezza di tutte le celle tranne l'ultima
            reduced_percentage = percentage / 2  # Puoi regolare il valore in base alle esigenze
            for i in range(0, n_columns):
                if i < n_columns - 1:
                    colWidths.append(doc.width * reduced_percentage)
                else:
                    remaining_width = doc.width - sum(colWidths)
                    colWidths.append(remaining_width)

        if table_type == 1:
            table = Table(data, colWidths=colWidths)
        else:
            table = Table(data, colWidths=colWidths, rowHeights=rowHeights)

        table.setStyle(TableStyle(style))
        return table



    def legend_table(self,doc,n_columns=8,percentage=0.09):
        data = [
                    ["DEVICES", "CANCELATION DEVIATION", "", "SESSION OUTCOME", "", "SESSION TYPE", "", "COLORS"],
                    ["FMS1", "CD1", "No Show user", "DCO", "Duty Carried Out", "STDA", "Stand Alone", "CANC SMC"],
                    ["FMS2", "CD2", "No show Instructor", "DNCO", "Duty Not Carried Out", "NET", "Linked Simulators", "RSLD"],
                    ["PTT1", "CD3", "Device failed during session", "CANC", "Mission Canceled", "", "", "All Other Outcomes"],
                    ["PTT2", "CD4", "Device not ready for training", "DPCO", "Duty Partially Carried Out", "", "", ""],
                    ["PTT3", "CD5", "2nd SIM not available for NETMODE", "RSLD", "Rescheduled", "", "", ""],
                    ["ULTD1", "CD6", "Device under service", "SMC", "Support Mission Canceled", "", "", ""],
                    ["ULTD2", "CD7", "Facility not ready","SDC", "Support Did Complete", "", "", ""],
                    ["", "CD8", "Software Development", "SDNC", "Support Did Not Complete", "", "", ""],
                    ["", "CD9", "Hardware Development", "ERR", "Error on DFS", "", "", ""],
                    ["", "CD10", "External Factors", "", "", "", "", ""],
                    ["", "CD11", "LVC failed during session", "", "", "", "", ""],
                ]
        style = TableStyle([
                                ('GRID', (0, 0), (-1, 0), 0.05, colors.lightgrey),
                                # (DEVICES)
                                ('SPAN', (0, 0), (0, 0)),
                                ('LINEBELOW', (0, 1), (0, -1), 0, colors.white),
                                
                                # (CANCELATION DEVIATION)
                                ('SPAN', (1, 0), (2, 0)),
                                ('GRID', (1, 1), (2, 11), 0.05, colors.lightgrey), #colors.lightgrey
                                
                                # (SESSION OUTCOME)
                                ('SPAN', (3, 0), (4, 0)),
                                ('GRID', (3, 1), (4, 9), 0.05, colors.lightgrey),
                                
                                # (SESSION TYPE)
                                ('SPAN', (5, 0), (6, 0)),
                                ('GRID', (5, 1), (6, 2), 0.05, colors.lightgrey),
                                
                                # (COLORS)
                                ('SPAN', (7, 0), (7, 0)),
                                ('BACKGROUND', (7, 1), (7, 1), "#dc3545"),
                                ('BACKGROUND', (7, 2), (7, 2), "#d3d3d3"), 
                                ('BACKGROUND', (7, 3), (7, 3), "#ffa500"),
                                ('LINEBELOW', (7, 1), (7, 3), 1, "#E9E9E9"),

                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),    
                                ('FONTNAME', (0,0), (-1,0), 'Times-Bold'),   # Header font
                                ('FONTNAME', (0,1), (-1,-1), 'Times-Roman'),
                                ('FONTSIZE', (0,0), (-1,-1), 5),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                ('TOPPADDING', (0, 0), (-1, -1), 5),
                                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),

                            ])
        table = Table(data,rowHeights=12)
        table.setStyle(style)
        return table
    
    def  session_summary_table(self,scheduled_time,scheduled_session,completed_time,completed_session):
        data = [
                    ["", "Sessions", "Time", ],
                    ["Total Scheduled", scheduled_session,scheduled_time,],
                    ["Total Completed", completed_session, completed_time],
                ]
        style = TableStyle([
                                ('GRID', (0, 0), (-1, -1), 0.05, colors.lightgrey),
                                ('SPAN', (0, 0), (0, 0)),
                                ('LINEBELOW', (0, 1), (0, -1), 0, colors.white),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),    
                                ('FONTNAME', (0,0), (-1,0), 'Times-Bold'),   # Header font
                                ('FONTNAME', (1,1), (-1,-1), 'Times-Roman'),
                                ('FONTNAME', (0,0), (0,-1), 'Times-Bold'),
                                ('FONTSIZE', (0,0), (-1,-1), 5),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                ('TOPPADDING', (0, 0), (-1, -1), 3),
                                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),

                            ])
        table = Table(data,rowHeights=12, hAlign="LEFT")
        table.setStyle(style)
        return table


    def table_generator_rtms(self,doc,data,style,percentage=0.12,n_columns=8,n_different_width=3): 
        colWidths=[]
        colWidths=[]
        for i in range(0,n_columns):
            if i>=n_columns-n_different_width:
                colWidths.append(doc.width*(1-((n_columns-1)*percentage)))
            else: colWidths.append(doc.width*percentage)
        table=Table(data, colWidths=colWidths)
        table.setStyle(TableStyle(style))
        return  table
    
    
    def set_up_today_rows(self, doc,data:list, header=['S/N','SESSION','STUDENT','INSTRUCTOR','SIM','LINKED SIM','ACT DURATION','OUTCOME','DEVIATION','COMMENT']):
        from reportlab.lib.units import mm
        from reportlab.platypus import Table, TableStyle
        from reportlab.lib import colors
        from reportlab.pdfgen import canvas
        today_row=[]
        today_row.append(header)
        row_backgrounds ={'gray':colors.lightgrey, 'white': colors.white, 'red':"#dc3545",'orange':colors.orange}   
        table_style_list=self.table_general_style(header_font_size=5)
        row_index=0
        for row in data:
            today_row.append(row[:len(row)-2])
            row_start = (0, row_index + 1)  # Row index + 1 to skip the header row
            row_end = (-1, row_index + 1)   # Row index + 1 to skip the header row
            #table_style_list.append(('BACKGROUND', row_start, row_end, row_backgrounds[row[len(row)-1]]))
            if row_index%2==0:
                table_style_list.append(('BACKGROUND', row_start, row_end, "#f5f5f5"))
            else: table_style_list.append(('BACKGROUND', row_start, row_end, colors.white))
            if row[len(row)-1] != "white":
                table_style_list.append(('BACKGROUND', (7,row_index + 1), (8,row_index + 1), row_backgrounds[row[len(row)-1]])) 
                
                           
            row_index+=1
        return self.table_generator(doc,today_row,table_style_list,percentage=0.09)
    

    def set_up_rows(self, doc,data:list, header=['S/N','SESSION','STUDENT','INSTRUCTOR','SIM','LINKED SIM','PLD DURATION','COMMENT'], table_type=1):
        tomorrow_row=[]
        tomorrow_row.append(header)
        table_style_list=self.table_general_style(header_font_size=5)   
        row_index=0
        if table_type==1:
            for row in data:
                row = ["" if x == "nan" else x for x in row]
                tomorrow_row.append(row[:len(row)])
                row_start = (0, row_index + 1)  # Row index + 1 to skip the header row
                row_end = (-1, row_index + 1)
                if row_index%2==0:
                    table_style_list.append(('BACKGROUND', row_start, row_end, "#f5f5f5"))
                else: table_style_list.append(('BACKGROUND', row_start, row_end, colors.white))
                row_index+=1
            return self.table_generator(doc,tomorrow_row,table_style_list,n_columns=8, table_type=1) 
        for row in data:
            tomorrow_row.append(row[:len(row)])
        return self.table_generator(doc,tomorrow_row,table_style_list,n_columns=8, table_type=1)
    
    def set_up_simulator_status(self,doc,data:list,header=['DEVICE','STATUS','ISSUE','DATE OF OCCURRANCE','ESTIMATED/RESOLVED TIME','COMMENT']):
        status_table=[]
        status_table.append(header)
        row_backgrounds ={'green':"#3cb371", 'red': "#dc3545",'orange':colors.orange}
        table_style_list=self.table_general_style(header_font_size=5,cells_padding=0.001)
        row_index=0
        for row in data:
            status_table.append(row[:len(row)-1])
            row_start = (1, row_index + 1)  # Row index + 1 to skip the header row
            row_end = (1, row_index + 1)   # Row index + 1 to skip the header row
            if row_index%2==0:
                    table_style_list.append(('BACKGROUND', (0, row_index + 1), (-1, row_index + 1), "#f5f5f5"))
            else: table_style_list.append(('BACKGROUND', (0, row_index + 1), (-1, row_index + 1), colors.white))
            row_index+=1
            table_style_list.append(('BACKGROUND', row_start, row_end, row_backgrounds[row[len(row)-1]]))
        return self.table_generator(doc,status_table,table_style_list,n_columns=6,percentage=0.24, table_type=0) 
    
    def set_up_cbt_sbt_status(self,doc,data:list,header=['DEVICE','STATUS','ISSUE','DATE OF OCCURRANCE','ESTIMATED/RESOLVED TIME','COMMENT']):
        status_table=[]
        status_table.append(header)
        row_backgrounds ={'green':"#3cb371", 'red': "#dc3545", 'gray':"#d3d3d3",'orange':colors.orange}
        table_style_list=self.table_general_style(header_font_size=5,cells_padding=0.001)
        row_index=0
        for row in data:
            status_table.append(row[:len(row)-1])
            row_start = (1, row_index + 1)  # Row index + 1 to skip the header row 
            row_end = (1, row_index + 1)   # Row index + 1 to skip the header row
            if row_index%2==0:
                    table_style_list.append(('BACKGROUND', (0, row_index + 1), (-1, row_index + 1), "#f5f5f5"))
            else: table_style_list.append(('BACKGROUND', (0, row_index + 1), (-1, row_index + 1), colors.white))
            table_style_list.append(('BACKGROUND', row_start, row_end, row_backgrounds[row[len(row)-1]]))
            row_index+=1
        table_style_list.append(('FONTNAME', (0,-1), (0,-1), 'Times-Italic'))
        return self.table_generator(doc,status_table,table_style_list,n_columns=6,percentage=0.24,table_type=0)

    def table_general_style(self,header_font_size=5,cells_padding=6): 
        table_style_list = [
          
        ('BACKGROUND', (0,0), (-1,0), colors.white),  # Header row
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),   # Header text color
        #('ALIGN', (0,0), (-1,-1), 'CENTER'),          # All cells alignment
        ('FONTNAME', (0,0), (-1,0), 'Times-Bold'),   # Header font
        ('FONTNAME', (0,1), (-1,-1), 'Times-Roman'),   # Content font
        ('FONTSIZE', (0,0), (-1,0), header_font_size),             # Header suitable to the table
        ('FONTSIZE', (0,1), (-1,-1), 4),             # Content font size
        ('BOTTOMPADDING', (0,0), (-1,0), 0.0002),        # Header bottom padding. Height of header
        ('TOPPADDING', (0, 0), (-1, -1), 7 ),
        ('BOTTOMPADDING', (0,1), (-1,-1), cells_padding),        # Content bottom padding. To edit
        #('GRID', (0,0), (-1,-1), 0.1,colors.black ),#"#E9E9E9"),
        ('INNERGRID', (0,0), (-1,-1), 0.21,colors.white ),
        ('BOX', (0,0), (-1,-1), 0.21, colors.white),
        ('WORDWRAP', (0,0), (-1,-1), 1),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),  # Optional: set left padding for content cells
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),  # Optional: set right padding for content cells
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Center text vertically
        ]
        # Set row heights
        
            
        return table_style_list
    
    

    def header_general_style(self):
        
        table_style_list = [
          
        ('BACKGROUND', (0,0), (-1,-1), "#789ECC"),  # Header row B4C6E7
        ('TEXTCOLOR', (0,0), (-1,-1), colors.black),   # Header text color
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),          # All cells alignment
        ('FONTNAME', (0,0), (-1,-1), 'Times-Bold'),   # Header font
        ('FONTSIZE', (0,1), (-1,-1), 8)          
                        ]
        return table_style_list


    def  text_generator(self,doc,text,text_type,font_size=12,fontName="Times-Bold"):
        styles = getSampleStyleSheet()
        
        #title=Paragraph(text,styles[text_type].clone(fontSize=font_size, fontName=fontName,name="custom"))
        base_style = styles[text_type]
        custom_style = ParagraphStyle(
                    name="custom",
                    parent=base_style,         
                    fontSize=font_size,
                    fontName=fontName,
                    spaceAfter=0,     
                    spaceBefore=0    
                )
        title = Paragraph(text, custom_style)
        return title


    def pdf_generator(self,doc,container=[]):
        try:
            #doc = SimpleDocTemplate(title,pagesize=landscape(letter))
            doc.build(container)
            return True
        except  Exception as pdf_err:
            print(pdf_err)
            return False
    
    
        
        
        

