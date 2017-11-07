import sys
import ctypes
import os.path
import csv
import time
import smtplib
from datetime import datetime
import pyodbc
import os
import argparse
from contextlib import closing
from apiclient.discovery import build
import httplib2
import xlwt
from oauth2client import client, file, tools
from googleapiclient.errors import HttpError
from pyodbc import IntegrityError
from ftplib import FTP, error_temp
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QCompleter, qApp, QWidget,\
    QLabel, QSizePolicy, QVBoxLayout, QTableWidgetItem, QFileDialog
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from ui import Ui_MainWindow, Ui_frmTblGoogle
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from PyQt5.QtGui import QMovie
import shutil
from _datetime import timedelta

class FTPWindow(QMainWindow, Ui_MainWindow):
    selected_file = None
    
    def __init__(self):
        super(FTPWindow, self).__init__()
        
        #connect to FTP server.
        self.connect_ftp()
        #setup UI from Qt Designer.
        self.setupUi(self)
        #display welcome from ftp server.
        #self.lbl_connection.setText(self.ftp.getwelcome())
        #get a list of files for James to choose from and put them in 
        #a combo box.
        files = self.get_file_list()
        self.cb_files.addItems(files)
        self.cb_upd_files.addItems(files)
        self.cb_cross_files.addItems(files)
        
        #setup a auto completer to choose campaign files from.
        campaigns = self.get_campaigns()
        camp_completer = QCompleter(campaigns)
        camp_completer.setCompletionMode(QCompleter.InlineCompletion)
        camp_completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
        camp_completer.setCaseSensitivity(Qt.CaseInsensitive)
        #we have to styple this here instead of Qt Designer, becuase for some reason there is no way
        #to style the popup using css. 
        camp_completer.popup().setStyleSheet("""selection-background-color: #ffaa00;
                background-color: QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #565656, stop: 0.1 #525252, stop: 0.5 #4e4e4e, stop: 0.9 #4a4a4a, stop: 1 #464646);
                border-style: solid;
                border: 1px solid #1e1e1e;
                border-radius: 5;""")       
        self.le_campaign.setCompleter(camp_completer)
        
        #connect the button signals with the appropriate slots.
        self.btn_dl_match.clicked.connect(self.btn_download_clicked)
        self.btn_match.clicked.connect(self.btn_match_clicked)
        
        self.btn_dl_update.clicked.connect(self.btn_download_clicked)
        
        self.btn_dl_cc.clicked.connect(self.btn_download_clicked)
        
        self.btn_google.clicked.connect(self.btn_google_clicked)
        
        self.actionExit.triggered.connect(qApp.instance().exit)
        
        max_date = self.get_last_google_import()
        self.start_dt = max_date + timedelta(days=1)
        self.end_dt = datetime.now() - timedelta(days=1)
        #self.de_ga_start.setDateTime(datetime.now() - timedelta(days=13))
        self.de_ga_start.setDateTime(self.start_dt)
        self.de_ga_end.setDateTime(self.end_dt)
        self.de_ga_start.dateChanged.connect(self.date_changed)
        self.de_ga_end.dateChanged.connect(self.date_changed)
        
        self.btn_view_ga_tbl.clicked.connect(self.view_table)

    def date_changed(self):
        de = self.sender()

        if de.objectName() == 'de_ga_start':
            self.start_dt = self.de_ga_start.date().toPyDate()
        else:
            self.end_dt = self.de_ga_end.date().toPyDate()
        
    def connect_ftp(self, company='ExactTarget'):
        try:
            if company == 'ExactTarget':
                #Connect to Exact Target FTP site.
                host = 'ftp.s4.exacttarget.com'
                user = '1005090'
                password = 'Rx7.k2.F'
                
            elif company == 'CrossCountry':
                #Connect to Cross Country FTP site.
                host = 'ftp.crosscountrycomputer.com'
                user = 'ccftpExt'
                password = 'ExT^23'
                
            self.ftp = FTP(host)
            self.ftp.login(user, password)
            #change the default directory for Exact Target.
            if company == 'ExactTarget':
                self.ftp.cwd('Export')            
        except error_temp as e:
            QMessageBox.critical(self,'Cannot Connect', 'Could not connect to the FTP server, please try again. \n %s' % str(e), QMessageBox.Ok)
        

    def get_file_list(self):
        #grabs the list of files in the currently connected ftp server.
        self.filenames = []
        try:
            self.ftp.retrlines('NLST', self.filenames.append)
        except error_temp as e:
            QMessageBox.critical(self,'Cannot Connect', 'Could not connect to the FTP server, please try again. \n %s' % str(e), QMessageBox.Ok)
        
        return self.filenames 
    
    def get_campaigns(self):
        #get a list of current campaigns for the database and return them in a list.
        connection = ConnectData()
        cursor = connection.connect_dmc()
        cursor.execute('select distinct IDName from dbo.ExternalSent order by IDName')
        ds = cursor.fetchall()
        
        ls_camp_name = []
        for row in ds:
            ls_camp_name.append(row[0])
        
        return ls_camp_name    
    
    def btn_download_clicked(self):
        try:
            #send a dummy operation to the ftp server to make sure we are still connected            
            self.ftp.voidcmd("NOOP")
        except IOError:
            #if not reconnect
            self.connect_ftp()
        
        #get the name of the button that sent us here.
        self.btn_clicked = self.sender().objectName()
         
        #start the download process based on what button sent us here.
        if self.btn_clicked == "btn_dl_match":
            #start downloaded with the selected file in the combo box.
            self.download_start(self.cb_files.currentText())
        elif self.btn_clicked == "btn_dl_update":
            #start the download with the file selected in the combo box.
            self.download_start(self.cb_upd_files.currentText())
        elif self.btn_clicked == "btn_dl_cc":
            #start the download with the file selected in the combo box.
            self.download_start(self.cb_cross_files.currentText())
            
        
    def download_start(self, selected_file):
        self.selected_file = selected_file
        
        #Call download thread
        self.download_file = DownloadFile(self.selected_file, self.ftp)
        #Start the thread
        self.download_file.start()
        #Connect the signals to the proper functions
        self.download_file.downloadProgress.connect(self.download_progress)
        self.download_file.download_finished.connect(self.download_finished)
        #####TO DO#####
        #Add error signal to download class.
        
    def download_finished(self):
        dir_name = 'download'
            
        if self.btn_clicked == "btn_dl_match":
            import_type = 'match'
            #set file name and path
            file_name = self.cb_files.currentText()
            file_path = os.path.join(dir_name, file_name)
            
            #set the progress bars to 100%, don't think we need this anymore, will check when James is done.
            self.pb_download_match.setRange(0,1)
            self.pb_download_match.setValue(1)
            
            #get row count of the sheet downloaded to track progress when uploading
            row_count = sum(1 for row in csv.reader(open(file_path))) 
            #set the range of the progress bar respectively
            self.pb_import_match.setRange(0, row_count)
            
        elif self.btn_clicked == "btn_dl_update":
            import_type = 'update'
            #set file name and path
            file_name = self.cb_upd_files.currentText()
            file_path = os.path.join(dir_name, file_name)            
            
            #set the download progress bar to 100%, don't think we need this anymore.
#             self.pb_download_update.setRange(0,1)
#             self.pb_download_update.setValue(1)
            
            #get row count and update progress bar, again.
            row_count = sum(1 for row in csv.reader(open(file_path)))
            self.pb_import_update.setRange(0, row_count)
            
        elif self.btn_clicked == "btn_dl_cc":
            import_type = 'ignore'
            #set file name and path
            file_name = self.cb_cross_files.currentText()
            file_path = os.path.join(dir_name, file_name)
            
            #Update progress bar, again.
            self.pb_download_cc.setRange(0,1)
            self.pb_download_cc.setValue(1)
            
            self.process_cc(file_path)
            
        self.import_start(file_path, import_type) 
        
    def import_start(self, file_path, import_type):
        #there has to be a better way to do this, but this starts the import
        #process and connect's the finished signal to the proper slot.
        self.import_csv = ImportCSV(file_path, import_type)
        self.import_csv.notify_progress.connect(self.on_progress)
        self.import_csv.start()
        #connect signal to function
        self.import_csv.import_finished.connect(self.import_finished)


#WTF?        
#         if self.btn_clicked == "btn_dl_match":
#             self.import_csv = ImportCSV(file_path, import_type)
#             self.import_csv.notify_progress.connect(self.on_progress)
#             self.import_csv.start()
#             #connect signal to function
#             self.import_csv.import_finished.connect(self.import_finished)
#         elif self.btn_clicked == "btn_dl_update":
#             self.import_csv = ImportCSV(file_path, import_type)
#             self.import_csv.notify_progress.connect(self.on_progress)
#             self.import_csv.start()
#             #connect finished signal to function
#             self.import_csv.import_finished.connect(self.import_finished)
            
        ###TO DO###
        #Add error signal to ImportCSV thread
        
    def import_finished(self):
        if self.btn_clicked == "btn_dl_match":
            #Set progress bars to 100%, don't think this needs done anymore.
            self.pb_import_match.setRange(0,1)
            self.pb_import_match.setValue(1)
            #Enable labels, line edits and buttons to run the next functions
            self.lbl_campaign.setEnabled(True)
            self.le_campaign.setEnabled(True)  
            self.btn_match.setEnabled(True)
            self.lbl_match.setEnabled(True)
            self.pb_compare.setEnabled(True)
        elif self.btn_clicked == "btn_dl_update":
            #set the match progress bar to pulse...
            self.pb_match_update.setRange(0,0)
            #and start the matching!
            self.match_subs_start()
        
    def on_progress(self, pro):
        #This is to track progress emitted from importing a csv file
        #into the database.
        if self.btn_clicked == "btn_dl_match":
            self.pb_import_match.setValue(pro)
        elif self.btn_clicked == "btn_dl_update":
            self.pb_import_update.setValue(pro)
        elif self.btn_clicked == "btn_dl_cc":
            self.pb_upload_cc.setValue(pro)
            
    def download_progress(self, pro):
        #For tracking progress emitted from downloading from FTP site, sets the in value 
        #for percent complete.
        if self.btn_clicked == "btn_dl_match":
            self.pb_download_match.setValue(pro)
        elif self.btn_clicked == "btn_dl_update":
            self.pb_download_update.setValue(pro)
        elif self.btn_clicked == "btn_dl_cc":
            self.pb_download_cc.setValue(pro)        
        
    def btn_match_clicked(self):
        #When "Start" button is clicked in "Find Subs..." tab.
        self.pb_compare.setRange(0,0)
        self.compare_start()

    def match_subs_start(self):
        #When "Match" button is clicked in "Find Subs..." tab.
        self.match = ExportNonMatched()
        self.match.send_export.connect(self.match_subs_finished)
        self.match.start()
        
    def match_subs_finished(self, ls_export, row_count):
        #After finding non matched email addresses.
        
        #Grab row count of the non matched records to return to James for records and confirmation.
        row_count = "{:,}".format(row_count)
        
        #Stop the pulsing on the "Finding Matches" progress bar.
        self.pb_match_update.setRange(0,1)
        self.pb_match_update.setValue(1)
        
        #Set up file name for writing 
        today = time.strftime("%m-%d-%y")
        base_name = 'NonMatchedEmailSubscriberDMCIDs_' + today
        file_suffix =  '.csv'
        
        #Set up full path of of file for writing and saving.
        file_name = os.path.join(r'\\esmserver\InkPixi\ONLINE MARKETING\Salesforce Marketing Cloud\EXPORTS\from Order Manager', base_name + file_suffix)   
        #Display to James how many records are going to be saved and where they are going to be saved to.
        reply = QMessageBox.question(self, 'Matches', row_count + ' Records will be exported to: \n' + file_name, QMessageBox.Ok| QMessageBox.Cancel)
        
        #If James is ok with the numbers...
        if reply == QMessageBox.Ok:
            #Create the .csv file and and write the list returned from non mataced to it
            try:
                with open(file_name, 'w', newline='') as output:
                    wr = csv.writer(output)
                    wr.writerow(['Email Address', 'DMCID'])
                    for val in ls_export:
                        wr.writerow([val[0], val[1]])
            except BaseException as e:
                #In case somebody has a duplicate sheet open, or there is a permission error, tell them about it.
                QMessageBox.information(self, 'Error', 'Something went wrong here, please try again or contact IT. \n ' + str(e), QMessageBox.Ok)          
        elif reply == QMessageBox.Cancel:
            ###TO DO###
            #Reset the tab.
            QMessageBox.information(self, 'Cancelled', 'Action cancelled')
            
        self.pb_download_update.setValue(0)
        self.pb_import_update.setValue(0)
        self.pb_match_update.setValue(0)
        
    def compare_start(self):
        #Start the thread to compare the downloaded file from "Find Subs..." tab
        #With the campaign ID that is selected from the campaign line edit
        self.compare = Compare(self.le_campaign.text())
        self.compare.send_results.connect(self.compare_finished)
        self.compare.start()

    def compare_finished(self, ls_results, row_count):
        #Gets list and rock count from comparing the downloaded file from "Find Subs..." tab
        #with selected campaign id. 
        row_count = "{:,}".format(row_count)
        
        #Stop the "Finding Matches" progress bar from pulsing.
        self.pb_compare.setRange(0,1)
        self.pb_compare.setValue(1)
        
        #Show James how many matches there were.
        QMessageBox.information(self, 'Matches', 'Amount of matches found: \n' + str(row_count),QMessageBox.Ok)
        
        #Set up initial file name, this is dumb I can do this better.
        base_name = self.le_campaign.text() + '-matched-subs'
        file_suffix =  '.csv'
        
        #Set the full path to create the file.
        file_name = os.path.join(r'\\esmserver\InkPixi\ONLINE MARKETING\Salesforce Marketing Cloud\ESM Postcard Groups', base_name + file_suffix)

        #Create and write to file from list returned from Compare()        
        with open(file_name, 'w', newline='') as output:
            wr = csv.writer(output)
            wr.writerow(['Campaign', 'DMCID', 'Email Address', 'Product Name'])
            for val in ls_results:
                wr.writerow([val[0], val[1], val[2], val[3]])
            output.close()
              
        #Reset setion to nothing so he can do more if he wants.
        self.le_campaign.setText('')                 
        
    def process_cc(self, file_path):
        #Set the processing progress bar to pulse while file manipulation and sorting is happening.
        self.pb_process_cc.setRange(0,0)
        #Set up the class
        self.process = ProcessCrossCountry(file_path)
        #Connect the results, and error signals to functions.
        self.process.process_results.connect(self.process_finished)
        self.process.process_error.connect(self.thread_error)
        #Fire it up.
        self.process.start()

    def process_finished(self, sortedList, rowCount, maxDt):
        #Export for CC tab...
        
        #For use further on down the road, from ProcessCrossCountry()
        self.ccRowCount = rowCount
        self.ccMaxDt = maxDt
        
        #Stop the pulsing of the progress bar by setting it complete.
        self.pb_process_cc.setRange(0,1)
        self.pb_process_cc.setValue(1)
        
        #Connect to Cross Country FTP site.
        self.connect_ftp('CrossCountry')
        
        #Set the upload progress bar 
        self.pb_upload_cc.setRange(0,100)
        
        #Set up the thread and connect the finished, error, and progress signals to proper functions.
        self.upload = UploadCrossCountry(sortedList, self.ftp)
        self.upload.uploadFinished.connect(self.upload_finished)
        self.upload.error.connect(self.thread_error)
        self.upload.progress.connect(self.on_progress)
        #Start the thread
        self.upload.start()
        
    def upload_finished(self):
        #Export CC tab...
        
        #Set up email thread with row count and max date from "Process Finished"
        self.sendEmail = SendEmail(self.ccRowCount, self.ccMaxDt)
        self.emailProgress = EmailProgress()
        
        #reconnect FTP to ExactTarget from Cross Country
        self.connect_ftp()
        
        #connect the finished signal to the proper function.
        self.sendEmail.emailSent.connect(self.cc_email_sent)
        
        #Start the email thread.
        self.sendEmail.start()
        
        #Show a little .gif so people know something is happening.
        self.emailProgress.show()
        
    def cc_email_sent(self):
        #Hide the .fig that was started in upload_finished.
        self.emailProgress.hide()
        
        #Show a box so people know that the email has been sent.
        QMessageBox.information(self, 'Email Sent', 'Email has been sent to Cross Country', QMessageBox.Ok)        
#         ok = QMessageBox.information(self, 'Email Sent', 'Email has been sent to Cross Country', QMessageBox.Ok)
#         if ok:
#             print('ok')
    
    def thread_error(self, txtError):
        #To show errors to user when something goes wrong...work in progress
        QMessageBox.critical(self, 'I got a bad feeling about this...', 'Something went wrong here... \n ' + txtError, QMessageBox.Ok)
    
    def btn_google_clicked(self):
        self.ga = ProcessGoogleData(self.start_dt, self.end_dt)
        self.ga.err.connect(self.thread_error)
        self.ga.finished.connect(self.google_finished)
        
        self.ga.start()
        self.pbar_google.setRange(0,0)
        
    def google_finished(self):
        self.pbar_google.setRange(0,1)
        self.pbar_google.setValue(1)
        
        self.view_table()
    
    def view_table(self):
        self.frmGoogle = GoogleTable()
        self.tblGoogle = self.frmGoogle.tblGoogle
        #fetch data thread
        self.gtd = GetGoogleTableData(self.start_dt, self.end_dt)
        self.gtd.err.connect(self.thread_error)
        self.gtd.finished.connect(self.goog_data_finished)
        
        self.pbar_google.setRange(0,0)
        self.gtd.start()
    
    def goog_data_finished(self, tbl_data):
        self.pbar_google.setRange(0,1)
        self.pbar_google.setValue(1)
        if tbl_data:
            self.goog_pop_table(tbl_data)
    
    def goog_pop_table(self, data):
        if data:
            self.tblGoogle.setRowCount(len(data))
            for i, row in enumerate(data):
                for j, col in enumerate(row):
                    item = QTableWidgetItem(str(col)) 
                    if item.text() == "None":
                        item.setText("")
                    self.tblGoogle.setItem(i, j, item)
        else:
            self.tblGoogle.setRowCount(1)
            item = QTableWidgetItem()
            item.setText("No Results")
            self.tblGoogle.setItem(0, 0, item)        
         
        self.tblGoogle.resizeColumnsToContents()
        self.tblGoogle.setSortingEnabled(True) 
        self.frmGoogle.show()        


    def export_google_tbl(self):
        self.tbl.export_table()
        
    def get_last_google_import(self):
        conn = ConnectData()
        cur = conn.connect_reporting()
        
        with closing(cur) as db:
            db.execute('select MAX(orderDate) orderDate from dbo.tblGoogleWebOrders')
            ds = db.fetchone()
        
        return ds[0]
        
class GoogleTable(QWidget, Ui_frmTblGoogle):
    
    def __init__(self):
        super(GoogleTable, self).__init__()   
        self.setupUi(self)
        
        self.btn_export.clicked.connect(self.export_table)
        
    def export_table(self):
        #choose path with the dialog
        path = QFileDialog.getSaveFileName(self, 'Choose Location', 'ga_export', ".xls(*.xls)")
        if path[0] != '':
            filename = os.path.abspath(path[0])
            #create instance of an excel workbook.
            wbk = xlwt.Workbook() 
            #create new sheet for table.
            sheet = wbk.add_sheet("SE vs. GA", cell_overwrite_ok=True)
            #create a list of the headers from the table to write to the sheet.
            lst_headers = []
            for h in range(self.tblGoogle.columnCount()):
                lst_headers.append(self.tblGoogle.horizontalHeaderItem(h).text())
            
            #write the headers to the work sheet.
            for i in range(len(lst_headers)):
                sheet.write(0, i, lst_headers[i])
            
            self.pop_sheet(sheet)
            wbk.save(filename)
        else:
            pass   
        
    def pop_sheet(self, sheet):
        for currentColumn in range(self.tblGoogle.columnCount()):
            for currentRow in range(self.tblGoogle.rowCount()):
                try:
                    txt = str(self.tblGoogle.item(currentRow, currentColumn).text())
                    sheet.write(currentRow +1, currentColumn, txt)
                except AttributeError:
                    pass  
        
class GetGoogleTableData(QThread):
    err = pyqtSignal(str)
    finished = pyqtSignal(list)
    
    def __init__(self, start_dt, end_dt):
        super(GetGoogleTableData, self).__init__()
        
        self.start_dt, self.end_dt = start_dt, end_dt
        
    def run(self):
        connection = ConnectData()
        cur = connection.connect_reporting()

        with closing(cur) as db:
            db.execute('EXEC dbo.uspGetGoogleAnalyticsVariance %s, %s' % (self.start_dt.strftime("'%m/%d/%Y'"), self.end_dt.strftime("'%m/%d/%Y'")))
            ds = db.fetchall()
            
            ls_qry = []
            for row in ds:
                ls_qry.append(row)
            
            self.finished.emit(ls_qry)
        
class ProcessGoogleData(QThread):
        err = pyqtSignal(str)
        finished = pyqtSignal()
        
        def __init__(self, start_dt, end_dt):
            super(ProcessGoogleData, self).__init__()
            
            self.start_dt, self.end_dt = start_dt, end_dt
            
        def run(self):
            # Define the auth scopes to request.
            scope = ['https://www.googleapis.com/auth/analytics.readonly']
            # Authenticate and construct service.
            service = self.get_service('analytics', 'v3', scope, 'client_secrets.json')
            profile = self.get_profile_id(service)

            results = self.get_data(service, profile, self.start_dt.strftime('%Y-%m-%d'), self.end_dt.strftime('%Y-%m-%d'))
            self.import_data(results)        
       
        def get_service(self, api_name, api_version, scope, client_secrets_path):
            """Get a service that communicates to a Google API.
            
            Args:
              api_name: string The name of the api to connect to.
              api_version: string The api version to connect to.
              scope: A list of strings representing the auth scopes to authorize for the
                connection.
              client_secrets_path: string A path to a valid client secrets file.
            
            Returns:
              A service that is connected to the specified API.
            """
            # Parse command-line arguments.
            parser = argparse.ArgumentParser(
                formatter_class=argparse.RawDescriptionHelpFormatter,
                parents=[tools.argparser])
            flags = parser.parse_args([])
            
            # Set up a Flow object to be used if we need to authenticate.
            flow = client.flow_from_clientsecrets(
                client_secrets_path, scope=scope,
                message=tools.message_if_missing(client_secrets_path))
        
            # Prepare credentials, and authorize HTTP object with them.
            # If the credentials don't exist or are invalid run through the native client
            # flow. The Storage object will ensure that if successful the good
            # credentials will get written back to a file.

            storage = file.Storage(api_name + '.dat')
            credentials = storage.get()
            http=httplib2.Http(ca_certs='cacerts.txt')
            if credentials is None or credentials.invalid:
                credentials = tools.run_flow(flow, storage, flags, http=http)
            http = credentials.authorize(http)
            
                # Build the service object.
            service = build(api_name, api_version, http=http)
                
            return service
        
        def get_profile_id(self, service):
            # Use the Analytics service object to get the first profile id.
            # Get a list of all Google Analytics accounts for the authorized user.
            accounts = service.management().accounts().list().execute()

            if accounts.get('items'):
                # Get the first Google Analytics account.
                account = accounts.get('items')[0].get('id')
                
                # Get a list of all the properties for the first account.
                properties = service.management().webproperties().list(
                    accountId=account).execute()
                
                if properties.get('items'):
                    # Get the first property id.
                    #property = properties.get('items')[0].get('id')
                    props = properties.get('items')
                    for i in range(len(props)):
                        if props[i].get('name') == 'InkPixi.com':
                            property =  props[i].get('id')
                    # Get a list of all views (profiles) for the first property.
                    profiles = service.management().profiles().list(accountId=account, webPropertyId=property).execute()
                 
                    if profiles.get('items'):
                        # return the first view (profile) id.
                        return profiles.get('items')[0].get('id')
                        
            return None
        
        def get_data(self, service, profileID, start_dt, end_dt):
            try:
                api_query = service.data().ga().get(
                        ids='ga:' + profileID,
                        start_date=start_dt,
                        end_date=end_dt,
                        metrics='ga:transactions, ga:transactionRevenue',
                        dimensions='ga:date',
                        sort='ga:date, ga:transactions',
                        #filters='ga:medium==organic',
                        max_results='5000')
            except TypeError as e:
                QMessageBox.critical(self, 'Error', str(e), QMessageBox.Ok)
            except HttpError as e:
                QMessageBox.critical(self, 'Error', str(e), QMessageBox.Ok)
            
            try:
                results = api_query.execute()
                return results
            except HttpError as err:
                self.err.emit(str(err))
            
              
        def import_data(self, results):
            conn = pyodbc.connect('DRIVER={SQL Server}; SERVER=SQLRPTSERVER\SQLREPORTS; DATABASE=OnlineMarketing; Trusted_Connection=yes')
            cur = conn.cursor()
            try:
                with closing(cur) as db:
                    columns = 'orderDate', 'orderCount', 'webRevenue'
                    query = 'insert into dbo.tblGoogleWebOrders({0}) values({1})'
                    query = query.format(','.join(columns), ','.join('?' * len(columns)))
                    try:
                        for row in results['rows']:
                            content = list(row[i] for i in range(len(columns)))
                            db.execute(query, content)
                        db.commit()
                    except IntegrityError as err:
                        self.err.emit('\n Some or all of the data you are trying to insert has already been imported, please check the dates. \n\n' +str(err))
            except TypeError as err:
                pass
            
            self.finished.emit()
                
class SendEmail(QThread):
        #QThread to send email after a file has been uploaded to Cross Country's FTP servers.
        emailSent = pyqtSignal()
        
        def __init__(self, rowCount = 0, lastOrder = None):
            super(QThread, self).__init__()
            self.rowCount = rowCount
            self.lastOrder = datetime.strftime(lastOrder, '%m/%d/%Y')
            
        def run(self):
            #Set up addresses for email.
            toEmail, ccEmail, fromEmail = ['ESM_Implementation@crosscountrycomputer.com', 'ainkles@crosscountrycomputer.com'], 'inkpixi.com@gmail.com', 'jwitmer@inkpixi.com'
            #toEmail, ccEmail, fromEmail = ['dhoy@aotees.com', 'hoy.davidj@gmail.com'], 'hoy.davidj@gmail.com', 'jwitmer@inkpixi.com'
            msg = MIMEMultipart('alternative')
            
            #set up and log into SMTP server.
            s = smtplib.SMTP('smtp.aotees.com', 25)
            s.login('problemsheets@aotees.com','esmtemp01')
            
            #Create message headers, these have nothing to do with the email being sent, more just place holders for the addresses.    
            msg['Subject'] = 'inkPixi files from ExactTarget are uploaded'
            msg['From'] = fromEmail
            msg['To'] = ', '.join(toEmail)
            msg['CC'] = ccEmail
            #create the body of the email
            body = 'You need to be viewing your email in html, please contact IT.'
            html = """\
            <html>
              <head></head>
              <body>
                <p>
                    The inkPixi files from ExactTarget have been uploaded to the Cross Country FTP site.<br>
                    Counts: """+str(self.rowCount)+"""<br>
                    Last Order Date: """+str(self.lastOrder)+"""<br><br>
                    Thanks, <br>
                    James
                </p>
              </body>
            </html>
            """
            #the toAll is how the email get's sent so we want to combine all the to's cc's and bcc's here in a list
            #SMTP doesn't care about cc or bcc it just want's a list of email's to send to. The cc, and bcc are taken
            #care of above in the head, for display in the email program.
            toAll = toEmail + [ccEmail]
            content = MIMEText(body, 'plain')
            html = MIMEText(html, 'html')
            
            #set up and send email.
            msg.attach(content)
            msg.attach(html)
            s.sendmail(fromEmail, toAll, msg.as_string())
            s.quit()     
            
            #Notify caller that thread is complete. 
            self.emailSent.emit()
            
class EmailProgress(QWidget):
    #Shows a .gif while email is being sent.
    def __init__(self):
        super(EmailProgress, self).__init__()
        
        #Sets the background to invisible.
        self.setAttribute(Qt.WA_TranslucentBackground)
        #Gets rid of frame around window.        
        self.setWindowFlags(Qt.FramelessWindowHint)
        
        #Create a movie object.
        self.movie = QMovie(self)
        
        #In case the gif is missing, display something, maybe.
        self.movieLabel = QLabel("No movie loaded")
        
        #I think there are some alignment issues here with this .gif, needs looked at, could
        #be a better way to do this.
        self.movieLabel.setAlignment(Qt.AlignAbsolute)
        self.movieLabel.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)

        #Could be a better way to do this with Qt Designer
        self.mainLayout = QVBoxLayout()
        self.mainLayout.addWidget(self.movieLabel)
        self.setLayout(self.mainLayout)

        self.resize(300, 300)
 
        self.movieLabel.setMovie(self.movie)
        self.movie.setFileName('image/email_progress.gif')
        
        #Start aligning the .gif on the screen, always needs fudged, there is 
        #probably a better way to do this. 
        pos = fw.pos()
 
        x = (pos.x() - ((self.width()/2) - (fw.width()/2))) + 100 
        y = (pos.y() - ((self.height()/2) - (fw.height()/2))) + 100
         
        self.setGeometry(x,y, 300, 300)
        #End aligning the .fig on the screen.
        
        #Start the .gif.
        self.movie.start()  
     
class UploadCrossCountry(QThread):
    uploadFinished = pyqtSignal()
    progress = pyqtSignal(int)
    error = pyqtSignal(str)
    sizeWritten = 0
    totalSize = 0
    lastShownPercent = 0
        
    def __init__(self, sortedList, ftp):
        super(QThread, self).__init__()
        self.sortedList = sortedList
        self.ftp = ftp
        
    def run(self):
        
        #Set up filename and path to create a file for uploading.
        today = time.strftime('%m%d%Y')
        upload_file = 'PIXI_DMCID_subs_' + today + '.csv'
        upload = os.path.join(r'upload', upload_file)  
        
        #Write to the file from the list create in ProcessCrossCountry()
        try:
            with open(upload, 'w', newline='') as up:
                wr = csv.writer(up)
                wr.writerow(['Email Address', 'DMCID', 'Last_order_date_HH', 'Last_Order_DB_HH'])
                for li in self.sortedList:
                    wr.writerow([li[0], li[1], li[2], li[3]])
                up.close()
            
            #Gets the total size of the file for tracking upload.
            self.totalSize = os.path.getsize(upload)
            
            #Upload the file, with a handler to track the progress.
            self.ftp.storbinary("STOR " + upload_file, open(upload, 'rb'), 1024, self.handle)
            
            #move to server for storage
            serverLocation = os.path.join(r'\\esmserver\inkpixi\ONLINE MARKETING\Salesforce Marketing Cloud\EXPORTS\from Exact Target\PIXI_DMCID_Subs', upload_file)
            shutil.move(upload, serverLocation)              
            
            #Tell caller thread finished successfully.
            self.uploadFinished.emit()
        except BaseException as e:
            #Tell the caller something went wrong, usually a permissions error because someone has the sheet open.
            self.error.emit(str(e))
            
    def handle(self, block):
        #Use to track the progress of the upload in percent.
        self.sizeWritten += 1024
        percentComplete = round((self.sizeWritten / self.totalSize) * 100)

        if (self.lastShownPercent != percentComplete):
            self.lastShownPercent = percentComplete
            #Send progress.
            self.progress.emit(percentComplete)
             
class ProcessCrossCountry(QThread):
    #Set up signals for notifying caller.
    process_results = pyqtSignal(list, int, datetime)
    process_error = pyqtSignal(str)
    
    def __init__(self, cross_file):
        super(QThread, self).__init__()
        self.cross_file = cross_file
        
    def run(self):
        today = time.strftime("%m/%d/%y")
        
        #Get row count from file that was downloaded from Exact Target.
        row_count = sum(1 for row in csv.reader(open(self.cross_file)))
        
        #Grab headers to write to our own file after sorting to grab date.        
        with open(self.cross_file, 'r') as f:
            reader = csv.reader(f)
            #get's header row in csv file
            headers = next(reader)
            f.close()
        
        #Open file and dump into list for sorting, probably a better way to do this.
        with open(self.cross_file, 'r') as f:
            reader = csv.reader(f)
            #Get rid of headers.
            next(reader)
            lsDt = []
            for dt in reader:
                if dt[2]:
                    #Format strings back into dates for sorting.
                    lsDt.append(datetime.strptime(dt[2], '%m/%d/%Y %I:%M %p'))
            f.close()
        maxDt = max(lsDt)
        
        #Set up file to write tracking information for marketing.
        fa = os.path.join(r'\\esmserver\inkpixi\ONLINE MARKETING\Salesforce Marketing Cloud\EXPORTS\from Exact Target', 'ExactTarget - Export Control Sheet.csv')
        try:
            with open(fa, 'a', newline='') as a:
                wr = csv.writer(a)
                #Append "Download Date", "Make Order Date", "Row Count", and headers for file to be uploaded.
                wr.writerow([today, maxDt, row_count, headers])
        except BaseException as e:
            #Opps...
            self.process_error.emit(str(e))
  
        #Create another list to create a sorted .csv file to upload to Cross County, found out we don't
        #Need to do this exactly, but it's here and a nice feature for now.   
        with open(self.cross_file, 'r') as sort:
            reader = csv.reader(sort)
            headers = next(reader)
            lsCsv = []
            for d in reader:
                #If there is no date, get a date, other wise errors were ensuing.
                dt = self.get_sort_date(d[2])
                lsCsv.append([d[0], d[1], dt, d[3]])
            
            #Sort the list by "Last Order Date" descending.
            sortedList = sorted(lsCsv, key= lambda row: datetime.strptime(row[2], '%m/%d/%Y %I:%M %p'), reverse = True)
            sort.close()        
        
        #Send back the list.
        self.process_results.emit(sortedList, row_count, maxDt)
        
    def get_sort_date(self, getDate):
        #If there is no date, assign a minimum date.
        minDate = '1/1/1900 12:00 AM'
        #Return either the date entered or the minimum date.
        return getDate or minDate
        
class DownloadFile(QThread):
    #Set up signals.
    download_finished = pyqtSignal()
    downloadProgress = pyqtSignal(int)
    sizeWritten = 0
    totalSize = 0
    lastShownPercent = 0
    
    def __init__(self, selected_file, ftp):
        super(QThread, self).__init__()
        self.selected_file = selected_file
        self.ftp = ftp
       
    def run(self):
        #Set up filename and path's
        local_filename = os.path.join(r'download', self.selected_file)
        self.lf = open(local_filename, "wb")
        #Get the size of the download before downloading.
        self.totalSize = self.ftp.size(self.selected_file)
        
        #Download file
        self.ftp.retrbinary("RETR %s" % self.selected_file, self.handle)
        self.lf.close() 
        
        self.download_finished.emit()
        
    def handle(self, block):
        #Track the progress of the download.
        self.sizeWritten += (4*1024)
        percentComplete = round((self.sizeWritten / self.totalSize) * 100)
        
        self.lf.write(block)       
        
        if (self.lastShownPercent != percentComplete):
            self.lastShownPercent = percentComplete
            #Send signal of the progress of the download.
            self.downloadProgress.emit(percentComplete)  
              
class ImportCSV(QThread):
    #For importing downloaded .csv's into the database.
    
    #Set up signals.
    import_finished = pyqtSignal()
    notify_progress = pyqtSignal(int)
    
    def __init__(self, file_path, import_type):
        super(QThread, self).__init__()
        self.file_path = file_path
        self.import_type = import_type
        
    def run(self):
        #There definitely has to be a better way to do this...sorry.
        if self.import_type == 'match':
            self.import_match()
        elif self.import_type == 'update':
            self.import_update()
            
    def import_match(self):
        #Set up connection to the database.
        connection = ConnectData()
        conn = connection.connect_dmc()
        #Empty the table that holds subscriber info.
        conn.execute('truncate table dbo.tblExactTargetSubscribers')
        conn.commit()
        
        cnt = 0
        with open(self.file_path, 'r') as f:
            reader = csv.reader(f)
            #get's rid of header row in csv file
            next(reader)
            #Column headings in sql table.
            columns = 'email', 'dmc_id'
            
            #Set up query.
            query = 'insert into dbo.tblExactTargetSubscribers({0}) values ({1})'
            query = query.format(','.join(columns), ','.join('?' * len(columns)))
            
            #Insert the data into the table.
            for data in reader:
                content = list(data[i] for i in range(2))
                conn.execute(query, content)
                #For tracking progress.
                cnt += 1
                #Notify caller of progress.
                self.notify_progress.emit(cnt)    
            conn.commit()
            f.close()
        #Tell caller finished. 
        self.import_finished.emit()
        
    def import_update(self):
        #Set up connection to database.
        conn = ConnectData().connect_dmc()
        #Empty table temp that holds subscriber info.
        conn.execute('truncate table dbo.tblExactTargetSubscriberDump')
        conn.commit()
        
        #Not sure whay this one starts at one, but too scared to change it right now.
        cnt = 1
        with open(self.file_path, 'r') as f:
            reader = csv.reader(f)
            #Gets rid of headers
            next(reader)
            
            #Column in table, only one.
            column = 'email'
            #Set up the query to accept the params.
            query = 'INSERT INTO dbo.tblExactTargetSubscriberDump({0}) values ({1})'
            query = query.format(column, ','.join('?'))
            #Insert contents of file into table.
            for data in reader:
                content = list(data[i] for i in range(1))
                conn.execute(query, content)
                #For tracking progress.
                cnt += 1
                #Send progress info.
                self.notify_progress.emit(cnt)
            conn.commit()
            f.close()
        #Say I'm done.
        self.import_finished.emit()
        
class ConnectData(object):
    #Only exists to connect to database.
    def connect_dmc(self):
        conn = pyodbc.connect('DRIVER={SQL Server}; SERVER=SQLSERVER; DATABASE=DMC; Trusted_Connection=yes')
        db = conn.cursor()
        
        return db    
    
    def connect_reporting(self):
        conn = pyodbc.connect('DRIVER={SQL Server}; SERVER=SQLRPTSERVER\SQLREPORTS; DATABASE=OnlineMarketing; Trusted_Connection=yes')
        db = conn.cursor()
        
        return db                
    
class Compare(QThread):
    send_results = pyqtSignal(list, int)
                           
    def __init__(self, campaign_id):
        super(QThread, self).__init__()
        self.campaign_id = campaign_id
        
    def run(self):
        #Set up connection to database.
        connection = ConnectData()
        conn = connection.connect_dmc()
        #Grab a list of all email address with the selected campaign id.
        conn.execute("""SELECT DISTINCT
                            es.IDName campaign,
                            es.DMCID,
                            et.email,
                            es.Product_Name
                        FROM 
                            dbo.ExternalSent es
                        JOIN
                            dbo.tblExactTargetSubscribers et ON CAST(et.dmc_id AS INT) = es.DMCID
                        WHERE
                            es.IDName = ?""", [self.campaign_id])
        ds = conn.fetchall()
        
        row_count = 0
        ls_matches = []
        for i in ds:
            ls_matches.append([i[0], i[1], i[2], i[3]])
            row_count += 1
        
        #Tell caller I'm finished.
        self.send_results.emit(ls_matches, row_count) 
        
class ExportNonMatched(QThread):
    #Returns list of people who aren't matched to a campaign.
    send_export = pyqtSignal(list, int)
    
    def __init__(self):
        super(QThread, self).__init__()
        
    def run(self):
        connection = ConnectData()
        conn = connection.connect_dmc()
        conn.execute("""SELECT DISTINCT
                            email, 
                            dmcid
                        FROM
                            (        
                                SELECT     
                                    es.Email, 
                                    acust.DMCID
                                FROM         
                                    dbo.tblExactTargetSubscriberDump es
                                INNER JOIN
                                    AlumniOriginals.dbo.Orders ao ON es.Email = ao.Email 
                                INNER JOIN
                                    AlumniOriginals.dbo.Customers acust ON ao.CustomerID = acust.CustomerID
                                WHERE acust.DMCID IS NOT NULL    
                                UNION ALL
                                SELECT     
                                    ts.Email, 
                                    acust3.DMCID
                                FROM         
                                    dbo.tblExactTargetSubscriberDump ts
                                INNER JOIN
                                    AlumniOriginalsArchives2003.dbo.Orders ao3 ON ts.Email = ao3.Email 
                                INNER JOIN
                                    AlumniOriginalsArchives2003.dbo.Customers acust3 ON ao3.CustomerID = acust3.CustomerID
                                WHERE acust3.DMCID IS NOT NULL    
                                UNION ALL
                                SELECT     
                                    ts.Email, 
                                    acust4.DMCID
                                FROM         
                                    dbo.tblExactTargetSubscriberDump ts
                                INNER JOIN
                                    AlumniOriginalsArchives2004.dbo.Orders ao4 ON ts.Email = ao4.Email 
                                INNER JOIN
                                    AlumniOriginalsArchives2004.dbo.Customers acust4 ON ao4.CustomerID = acust4.CustomerID
                                WHERE acust4.DMCID IS NOT NULL    
                                UNION ALL
                                SELECT     
                                    ts.Email, 
                                    rcust.DMCID
                                FROM         
                                    dbo.tblExactTargetSubscriberDump ts
                                INNER JOIN
                                    ESMRetail.dbo.Orders ro ON ts.Email = ro.Email 
                                INNER JOIN
                                    ESMRetail.dbo.Customers rcust ON ro.CustomerID = rcust.CustomerID
                                WHERE rcust.DMCID IS NOT NULL    
                                UNION ALL
                                SELECT     
                                    ts.Email, 
                                    ipcust.DMCID
                                FROM         
                                    dbo.tblExactTargetSubscriberDump ts
                                INNER JOIN
                                    InkPixi.dbo.Orders ipo ON ts.Email = ipo.Email 
                                INNER JOIN
                                    InkPixi.dbo.Customers ipcust ON ipo.CustomerID = ipcust.CustomerID
                                WHERE ipcust.DMCID IS NOT NULL
                        
                            ) qry
                        WHERE
                            qry.dmcid IS NOT NULL
                        ORDER BY email""")
        ds = conn.fetchall()
        
        row_count = 0
        
        ls_non_match= []
        for i in ds:
            ls_non_match.append([i[0], i[1]])
            row_count += 1
            
        self.send_export.emit(ls_non_match, row_count) 
                
if __name__ == "__main__":
    myappid = 'Exact Target Email Compare'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid) 
    
    app = QApplication(sys.argv)
    app.setStyle("plastique")
    fw = FTPWindow()
    fw.show()
    
    sys.exit(app.exec_())