import os
import sys
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)
from generate import *

import datetime
from reportlab.lib.units import inch
from reportlab.platypus import Spacer, PageBreak, Image, PageTemplate, Frame, BaseDocTemplate, Paragraph
from io import BytesIO
#from PIL import Image 
#from reportlab.pdfgen import canvas
from email_utils import *
try:
    from dotenv import dotenv_values, load_dotenv
except:
    from dotenv import main 





path="//192.168.1.125/gbts/Create_PDF_Report/"  

#path= os.path.dirname(__file__)
try:
    env=dotenv_values(path+'env.env')
    logbookSH_path=env.get('source_file')
    file_url_lgbksh=env.get('file_url_lgbksh')
    dest_file=env.get('dest_file')
    log_file=env.get('log_file')
    report_path=env.get('report_path')
    sbtcbt_file=env.get('sbtcbt_file')
    logo_ajt=env.get('logo_path')
    rtms_log=env.get('rtms_log')
    file_url_rtms=env.get('file_url_rtms')
    file_url_daily=env.get('file_url_daily')
    site_url=env.get('site_url')
    new_daily=env.get('new_daily')
    username=env.get('username')
    password=env.get('password')
    mpds_stations_url=env.get('mpds_stations_url')
    file_mpds_stations=env.get('file_mpds_stations')
    client_id=env.get('client_id')
    shared_secret=env.get('shared_secret')
    tenant_id=env.get('tenant_id')
except: 
    env=main.load_dotenv(path+'env.env')
    logbookSH_path=os.getenv('source_file')
    file_url_lgbksh=os.getenv('file_url_lgbksh')
    dest_file=os.getenv('dest_file')
    log_file=os.getenv('log_file')
    report_path=os.getenv('report_path')
    sbtcbt_file=os.getenv('sbtcbt_file')
    logo_ajt=os.getenv('logo_path')
    rtms_log=os.getenv('rtms_log')
    file_url_rtms=os.getenv('file_url_rtms')
    file_url_daily=os.getenv('file_url_daily')
    site_url=os.getenv('site_url')
    new_daily=os.getenv('new_daily')
    username=os.getenv('username')
    password=os.getenv('password')
    mpds_stations_url=os.getenv('mpds_stations_url')
    file_mpds_stations=os.getenv('file_mpds_stations')
    client_id=os.getenv('client_id')
    shared_secret=os.getenv('shared_secret')
    tenant_id=os.getenv('tenant_id')

def add_logo(canvas, doc):
    canvas.saveState()
    logo_width = 100
    logo_height = 36
    x = doc.width - 50
    y = doc.height  +80
    logo = Image(logo_ajt, width=logo_width, height=logo_height)
    logo.drawOn(canvas, x, y)
    canvas.restoreState()



log=log_definition()
def create(today="",tomorrow=""):
    try:
        if len(sys.argv) > 2:
            if sys.argv[1] == '--today':
                today = sys.argv[2]
            if sys.argv[3] == '--tomorrow':
                tomorrow = sys.argv[4]
            if day_off(today).working_day()[0]=="":
                print("holidays")
                return
        elif day_off().working_day()[3]==True: return #During the weekend the report must not be generated
        else:
            today=day_off().working_day()[0]
            tomorrow=day_off().working_day()[1]
            #today="2025-10-20"
            #tomorrow="2025-10-21"
        if today=="":
            print("holidays")
            return
        log.open_log_file(log_file+"log.log","Script is checking Logbook SH")
        now = datetime.now()
        hour = now.time().hour
        minute = now.time().minute
        weekday = now.weekday()
        #Start check
        if not os.path.isfile(logbookSH_path): #file not found
            log.error_message(0,log_file)
            return 
        if log.is_report_uploaded(dest_file,today) == True: #Report already loaded
            log.error_message(3,log_file)
            return
       

        #----START TO UPLOAD THE DOCUMENTS FROM SHAREPOINT TO SHARED FOLDER----
        data_container=data()
        data_container.download_from_sharepoint(site_url,file_url_rtms,rtms_log,client_id,shared_secret,tenant_id)#upload the rtms logbook
        data_container.download_from_sharepoint(site_url,file_url_daily,new_daily,client_id,shared_secret,tenant_id)#the logbook event issue
        data_container.download_from_sharepoint(site_url,file_url_lgbksh,logbookSH_path,client_id,shared_secret,tenant_id)#Logbook SH
        data_container.download_from_sharepoint(site_url,mpds_stations_url,file_mpds_stations,client_id,shared_secret,tenant_id)#MPDS Stations
        
        

        lgbk=data_container.load_file(logbookSH_path,"Log Book")
        if lgbk.empty:
            log.error_message(2,log_file)
            if (((0 <= weekday <= 3) and hour == 19 and minute == 00) or (weekday == 4 and hour == 17 and minute == 00)):
                sender()
            return 
        if log.analyze_today_rows(lgbk,log_file,today)==0:
            if (((0 <= weekday <= 3) and hour == 19 and minute == 00) or (weekday == 4 and hour == 17 and minute == 00)):
                sender()
            return         
        #Stop check
        

        #Start to read the data
        today_training,pld_today,act,scheduled_sessions,executed_sessions=data_container.read_today_data(lgbk,today)
        if len(today_training)==0:
            log.error_message(2,log_file)
            if (((0 <= weekday <= 3) and hour == 19 and minute == 00) or (weekday == 4 and hour == 17 and minute == 00)):
                sender()
            return
        tomorrow_training, pld_tomorrow=data_container.read_tomorrow_data(lgbk,tomorrow)
        lgbk=data_container.load_file(logbookSH_path,"SIM STATUS")
        sim_status=data_container.read_sim_status(lgbk)
        #lgbk=data_container.load_file(logbookSH_path,"RTMS LVC MPDS STATUS") #STATUS from Logbook SH
        #rtms_lvc_status=data_container.read_sim_status(lgbk) #STATUS from Logbook SH
        
        
        mpds=data_container.load_file(file_mpds_stations,"MPDS STATIONS")
        cbt_sbt_lgbk=data_container.load_file(new_daily,'Readiness')
        os.remove(new_daily)
        os.remove(file_mpds_stations)

        sbt_status,cbt_status =data_container.read_cbt_sbt_status_new(cbt_sbt_lgbk,today)
        #sbt_status,cbt_status =data_container.read_cbt_sbt_status(cbt_sbt_lgbk)
        lgbk=data_container.load_file(rtms_log,"SIM STATUS") #STATUS from MPDS Logbook
        rtms_lvc_status=data_container.read_sim_status(lgbk)#STATUS from MPDS Logbook
        mpds_status=data_container.read_mpds_status(mpds)#MPDS Status
        rtms=data_container.load_file(rtms_log,"RTMS LOGBOOK") 
        rtms=data_container.read_rtms_data(rtms,today) 
        #Stop to read the data

        #Start to generate the pdf
        spacer = Spacer(1, 0.2 * inch)
        spacer1 = Spacer(1, 0.08 * inch)
        spacer2 = Spacer(1, 0.01 * inch)
        Spacer(1,0)

        doc = BaseDocTemplate(report_path+"GBTS_Daily_Report_"+str(today)+".pdf",pagesize=landscape(letter))
        pdf=pdf_dev()
        
        header=Table([['Simulators']],colWidths=(doc.width/10.0)*10) 
        header.setStyle(TableStyle(pdf.header_general_style()))
        header_sbt=Table([['SBT']],colWidths=(doc.width/10.0)*10) 
        header_sbt.setStyle(TableStyle(pdf.header_general_style()))
        header_cbt=Table([['CBT']],colWidths=(doc.width/10.0)*10) 
        header_cbt.setStyle(TableStyle(pdf.header_general_style()))
        header_rtsm_lvc= Table([['LVC/RTMS']],colWidths=(doc.width/10.0)*10) 
        header_rtsm_lvc.setStyle(TableStyle(pdf.header_general_style()))
        
        header_mpds=Table([['MPDS']],colWidths=(doc.width/10.0)*10) 
        header_mpds.setStyle(TableStyle(pdf.header_general_style()))

        header_rtms=Table([['RTMS']],colWidths=(doc.width/10.0)*10) 
        header_rtms.setStyle(TableStyle(pdf.header_general_style()))
        header_legend= Table([['Legend']],colWidths=(doc.width/10.0)*5.7) 
        header_legend.setStyle(TableStyle(pdf.header_general_style()))

        today_training_table=pdf.set_up_today_rows(doc,today_training)
        tomorrow_training_table=pdf.set_up_rows(doc,tomorrow_training)
        sim_status_table=pdf.set_up_simulator_status(doc,sim_status)
        sbt_status_table=pdf.set_up_cbt_sbt_status(doc,sbt_status)
        cbt_status_table=pdf.set_up_cbt_sbt_status(doc,cbt_status)
        mpds_status_table=pdf.set_up_cbt_sbt_status(doc,mpds_status)

        rtms_lvc_status_table=pdf.set_up_simulator_status(doc,rtms_lvc_status)
        if len(rtms)!=0:
            rtms_today=pdf.set_up_rows(doc,rtms,header=['S/N','C/S','A/C','POD ID','Radio CH','RTMS IP Comments','A/C IP comments','MPDS Comments'],table_type=1) #RTMS data present
        else: rtms_today=pdf.text_generator(doc,text="No Data Available",text_type='BodyText',font_size=6,fontName="Times-Roman")  
        legend=pdf.legend_table(doc)
        
        document_title=pdf.text_generator(doc,text="GBTS Daily Report "+str(today),text_type="Title",font_size=22,fontName="Times-Bold")
        performed_training=pdf.text_generator(doc,text="GBTS Performed Training",text_type='Heading1',font_size=10)
        activity_request=pdf.text_generator(doc,text="GBTS Activity Request "+str(tomorrow),text_type='Heading1',font_size=10)
        status_overview=pdf.text_generator(doc,text="GBTS Status Overview",text_type='Heading1',font_size=10)
        today_scheduled=pdf.text_generator(doc,text="Total Scheduled &nbsp;&nbsp;&nbsp;&nbsp;Time: "+str(pld_today)+"&nbsp;&nbsp;&nbsp;Sessions: "+str(scheduled_sessions),text_type='BodyText',font_size=8,fontName="Times-Roman")
        tomorrow_scheduled=pdf.text_generator(doc,text="Total Scheduled Time: "+str(pld_tomorrow),text_type='BodyText',font_size=6,fontName="Times-Roman")
       
        today_act=pdf.text_generator(doc,text="Total Completed &nbsp;&nbsp;&nbsp;Time: "+str(act)+"&nbsp;&nbsp;&nbsp;Sessions: "+str(executed_sessions),text_type='BodyText',font_size=8,fontName="Times-Roman")
        
        rtms_performed_training=pdf.text_generator(doc,text="RTMS Performed Training",text_type='Heading1',font_size=10)
        today_time_sessions=pdf.session_summary_table(pld_today,scheduled_sessions,act,executed_sessions)
        
        frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='frame')
        template=PageTemplate(id="GBTS_Daily_Report",frames=[frame])
        doc.addPageTemplates([template])
        template.beforeDrawPage=add_logo
        

       
        
        pdf.pdf_generator(doc,container=[                               document_title,spacer,
                                                                        status_overview,
                                                                        header,sim_status_table,spacer,
                                                                        header_sbt,sbt_status_table,spacer,
                                                                        header_cbt,cbt_status_table,spacer,
                                                                        header_rtsm_lvc,rtms_lvc_status_table,spacer,
                                                                        header_mpds,mpds_status_table,PageBreak(), 
                                                                        performed_training,#spacer1,
                                                                        #header,today_training_table,spacer1,today_scheduled, today_act,PageBreak(),
                                                                        header,today_training_table,spacer1,today_time_sessions, PageBreak(),
                                                                        header_legend,legend,PageBreak(),
                                                                        activity_request,spacer1,
                                                                        header,tomorrow_training_table, spacer1,tomorrow_scheduled,PageBreak(),
                                                                        rtms_performed_training,spacer1,
                                                                        header_rtms,rtms_today                                                                       
                                                                        ])
        
        
            
        print("Report generated and saved to "+dest_file+"GBTS_Daily_Report_"+str(today)+".pdf")
        with open(dest_file+"isloaded.txt", 'w') as file:#create again the file isloaded.txt
            file.write(today)
        return True
        #Stop to generate the pdf
    except Exception as generationErr:
        print(generationErr)
        return False



if __name__=='__main__':
    create() 
    
