import sys
import ctypes
import os.path
import csv
import time
import smtplib
import shutil
import posixpath
import pyodbc
import pysftp
import _cffi_backend #this is needed to run pysftp\paramiko when compiled.
from contextlib import closing
from datetime import datetime, timedelta
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, qApp, QWidget,\
    QLabel, QSizePolicy, QVBoxLayout, QTableWidgetItem, QFileDialog
from PyQt5 import QtNetwork, QtWebKit, QtPrintSupport # these are needed for the qwebview for some reason.
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QUrl
from ui import Ui_MainWindow
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from PyQt5.QtGui import QMovie

class FTPWindow(QMainWindow, Ui_MainWindow):
    selected_file = None
    
    def __init__(self):
        super(FTPWindow, self).__init__()

        #setup UI from Qt Designer.
        self.setupUi(self)        

        ### Start of new button set for listrak revamp ###
        
        #opens file dialog to choose file
        self.btn_open_dmc.clicked.connect(self.btn_open_dmc_clicked)
        self.btn_open_cc.clicked.connect(self.btn_open_cc_clicked)
        
        #starts the import into the database of the chosen file from above (btn_match)
        self.btn_import.clicked.connect(self.btn_import_clicked)
        self.btn_process_cc.clicked.connect(self.btn_process_cc_clicked)
        #add Daves autoscale tool to the webview.
        self.wv_autoscale.load(QUrl('http://inkpixi:4000/autoscale.php'))
        
        self.actionExit.triggered.connect(qApp.instance().exit)
        
        self.cb_garments.addItem('-- Select --')
        garments = self.get_garments()
        self.cb_garments.addItems(garments)
        #self.cb_garments.currentIndexChanged.connect(self.get_top_fifty_start)
        self.btn_top_fifty.clicked.connect(self.get_top_fifty_start)
        
        now = datetime.now()
        from_now = now - timedelta(days=90)

        self.de_from.setDate(from_now)
        self.de_to.setDate(now)
        
    def btn_open_dmc_clicked(self):
        file = self.choose_file()
        self.le_match.setText(file)
        self.import_type = 'update_dmcid'
        
    def btn_open_cc_clicked(self):
        file = self.choose_file()
        self.le_cross_country.setText(file)
        self.import_type = 'cross_country'
        
    def choose_file(self):
        dir_path = os.path.join('C:\\Users', os.getlogin(), 'Documents')
        #this returns a tuple and I only want the path so by using wtf I can return just the path in string form.
        file, wtf = QFileDialog.getOpenFileName(self, 'Choose file.', dir_path, "Flat Files (*.csv)")
        
        filename = os.path.normpath(file)
        
        return filename    
    
    def choose_save_file(self):
        dir_path = os.path.join('C:\\Users', os.getlogin(), 'Documents')
        filename, wtf = QFileDialog.getSaveFileName(self, 'Save Export', dir_path, 'Flat Files (*.csv)', 'test.csv')
        
        return filename
    
    def btn_import_clicked(self):
        #set file name and path
        file_path = self.le_match.text()
        
        if file_path:
            row_count = sum(1 for row in csv.reader(open(file_path)))
            
            self.pb_import_update.setRange(0, row_count)
            self.import_start(file_path, self.import_type)            
        else:
            QMessageBox.information(self, 'Choose File', 'Please choose a file to import')

    def btn_process_cc_clicked(self):
        file_path = self.le_cross_country.text()
        self.process_cc(file_path) 

    def postcard_compare(self, date):
        date.toPyDate()  
    
    def get_garments(self):
        #gets garments for top fifty combo box.
            con = ConnectData()
            cur = con.connect_import_export()
            cur.execute("""SELECT 
                                [Garment Code],
                                [Garment Description]
                            FROM 
                                dbo.GarmentCodes
                            WHERE
                                IP = 1
                            ORDER BY [Garment Code]""")
            ds = cur.fetchall()
            
            ls_garments = []
            for row in ds:
                ls_garments.append(row[0] + ' - ' + row[1])
                
            return ls_garments
  
    def get_top_fifty_start(self):
        sku_garment = self.cb_garments.currentText()
        sku = sku_garment[:3]
        
        self.get_top_fifty = GetTopFifty(sku, self.de_from.text(), self.de_to.text())
        self.get_top_fifty.start()
        
        self.progress = EmailProgress()
        self.progress.show()
        self.get_top_fifty.top_fifty.connect(self.get_top_fifty_finished)
    
    def get_top_fifty_finished(self, lst_top_fifty):
        self.progress.hide()
        if lst_top_fifty:
            self.tbl_top_fifty.setRowCount(len(lst_top_fifty))
            for i, row in enumerate(lst_top_fifty):
                for j, col in enumerate(row):
                    if j in [2, 3]:
                        item = QTableWidgetItem('{0:,}'.format(col))
                    elif j == 4:
                        item = QTableWidgetItem('${0:,.2f}'.format(col))
                    else:
                        item = QTableWidgetItem(str(col))
                    if item.text() == 'None':
                        item.setText(None)
                    #add the items from the database                                               
                    self.tbl_top_fifty.setItem(i, j, item)
                    
            self.tbl_top_fifty.show()
            self.tbl_top_fifty.resizeColumnsToContents()
        else:
            QMessageBox.information(self, 'No Results Found', 'Unable to find any results', QMessageBox.Ok)
       
        
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
        #I think I found a bit better of a way...
        #process and connect's the finished signal to the proper slot.
        
        if self.import_type == 'update_dmcid':
            self.import_csv = ImportCSV(file_path, import_type)
            self.import_csv.notify_progress.connect(self.on_progress)
            self.import_csv.start()
            #connect signal to function
            self.import_csv.import_finished.connect(self.import_finished)


        ###TO DO###
        #Add error signal to ImportCSV thread
        
    def import_finished(self):
        if self.import_type == "btn_match":
            #Set progress bars to 100%, don't think this needs done anymore.
            self.pb_import_match.setRange(0,1)
            self.pb_import_match.setValue(1)
            #Enable labels, line edits and buttons to run the next functions
            self.lbl_campaign.setEnabled(True)
            self.le_campaign.setEnabled(True)  
            self.btn_match.setEnabled(True)
            self.lbl_match.setEnabled(True)
            self.pb_match_camp.setEnabled(True)
        elif self.import_type == "update_dmcid":
            #set the match progress bar to pulse...
            self.pb_match_update.setRange(0,0)
            #and start the matching!
            self.match_dmc_start()
        
    def on_progress(self, pro):
        #This is to track progress emitted from importing a csv file
        #into the database.
        if self.import_type == "btn_dl_match":
            self.pb_import_match.setValue(pro)
        elif self.import_type == "update_dmcid":
            self.pb_import_update.setValue(pro)
        elif self.import_type == "cross_country":
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
        

    def match_dmc_start(self):
        #When "Match" button is clicked in "Find Subs..." tab.
        self.match_dmc_id = MatchDMCID()
        self.match_dmc_id.match_finished.connect(self.match_dmc_finished)
        self.match_dmc_id.start()
        
    def match_dmc_finished(self):
        self.pb_match_update.setRange(0,1)
        self.pb_match_update.setValue(1)
        #Enable labels, line edits and buttons to run the next functions
  
        self.pb_match_camp.setEnabled(True)

        #After finding non matched email addresses.
        
        #Grab row count of the non matched records to return to James for records and confirmation.
#         row_count = "{:,}".format(row_count)
#         
#         #Stop the pulsing on the "Finding Matches" progress bar.
#         self.pb_match_update.setRange(0,1)
#         self.pb_match_update.setValue(1)
#         
#         #Set up file name for writing 
#         today = time.strftime("%m-%d-%y")
#         base_name = 'NonMatchedEmailSubscriberDMCIDs_' + today
#         file_suffix =  '.csv'
#         
#         #Set up full path of of file for writing and saving.
#         file_name = os.path.join(r'\\esmserver\InkPixi\ONLINE_MARKETING\Salesforce Marketing Cloud\EXPORTS\from Order Manager', base_name + file_suffix)   
#         #Display to James how many records are going to be saved and where they are going to be saved to.
#         reply = QMessageBox.question(self, 'Matches', row_count + ' Records will be exported to: \n' + file_name, QMessageBox.Ok| QMessageBox.Cancel)
#         
#         #If James is ok with the numbers...
#         if reply == QMessageBox.Ok:
#             #Create the .csv file and and write the list returned from non mataced to it
#             try:
#                 with open(file_name, 'w', newline='') as output:
#                     wr = csv.writer(output)
#                     wr.writerow(['Email Address', 'DMCID'])
#                     for val in ls_export:
#                         wr.writerow([val[0], val[1]])
#             except BaseException as e:
#                 #In case somebody has a duplicate sheet open, or there is a permission error, tell them about it.
#                 QMessageBox.information(self, 'Error', 'Something went wrong here, please try again or contact IT. \n ' + str(e), QMessageBox.Ok)          
#         elif reply == QMessageBox.Cancel:
#             ###TO DO###
#             #Reset the tab.
#             QMessageBox.information(self, 'Cancelled', 'Action cancelled')
#             

        
        self.match_camps_start()
        
    def match_camps_start(self):
        #Start the thread to compare the downloaded file from "Find Subs..." tab
        #With the campaign ID that is selected from the campaign line edit
        self.compare = MatchCampaigns()
        self.compare.send_results.connect(self.match_camps_finished)
        self.pb_match_camp.setRange(0,0)
        self.compare.start()

    def match_camps_finished(self, ls_results, row_count):
        #Gets list and rock count from comparing the downloaded file from "Find Subs..." tab
        #with selected campaign id. 
        row_count = "{:,}".format(row_count)
        
        #Stop the "Finding Matches" progress bar from pulsing.
        self.pb_match_camp.setRange(0,1)
        self.pb_match_camp.setValue(1)
        
        #Show James how many matches there were.
        QMessageBox.information(self, 'Matches', 'Amount of records found: \n' + str(row_count) + '\nClick OK to save export.',QMessageBox.Ok)
        
        file_name = self.choose_save_file()
        
        #Create and write to file from list returned from Compare()  
        try:      
            self.save_export(file_name, ls_results)
        except PermissionError:
            ok = QMessageBox.warning(self, 'Close File', 'Please make sure that the file is closed and you have permission to it, if the file is open please close it now and then click ok.', QMessageBox.Ok | QMessageBox.Cancel)
            
            if ok == QMessageBox.Ok:
                file_name = self.choose_save_file()
                #there still semms to be an issue please call IT or something along those lines.
                try:
                    self.save_export(file_name, ls_results)
                except BaseException as e:
                    QMessageBox.warning(self, 'Something still seems wrong.', 'Something still seems to be wrong with the file you are saving too. \n' + str(e) + '\n Please call IT.')
            else:
                yes_no = QMessageBox.warning(self, 'Let\'s call the whole thing off', 'Are you sure you want to abort the import and matching?', QMessageBox.Yes | QMessageBox.No)
                if yes_no == QMessageBox.No:
                    file_name = self.choose_save_file()
                    try:
                        self.save_export(file_name, ls_results)
                    except BaseException as e:
                        QMessageBox.warning(self, 'Something still seems wrong.', 'Something still seems to be wrong with the file you are saving too. \n' + str(e) + '\n Please call IT.')
                else:
                    sys.exit()
            
        self.pb_import_update.setValue(0)
        self.pb_match_update.setValue(0)
        self.pb_match_camp.setValue(0)
        
        QMessageBox.information(self, 'File Saved', 'File has been saved to ' + file_name)
        
    def save_export(self, file_name, ls_results):
        with open(file_name, 'w', newline='') as output:
            wr = csv.writer(output)
            wr.writerow(['Email Address', 'Subscriber_info\DMCID', 'Subscriber_info\Last_DM_group'])
            for val in ls_results:
                wr.writerow([val[0], val[1], val[2]])
            output.close()        
        
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
        #maybe...
        #ftp = self.connect_cc_ftp()
        
        #or try...
        
        #For use further on down the road, from ProcessCrossCountry()
        self.ccRowCount = rowCount
        self.ccMaxDt = maxDt
        
        #Stop the pulsing of the progress bar by setting it complete.
        self.pb_process_cc.setRange(0,1)
        self.pb_process_cc.setValue(1)
        
        #Connect to Cross Country FTP site.
        #self.connect_ftp('CrossCountry')
        
        #Set the upload progress bar 
        self.pb_upload_cc.setRange(0,100)
        
        #Set up the thread and connect the finished, error, and progress signals to proper functions.
        self.upload = UploadCrossCountry(sortedList)
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
        ss = os.path.join(r'\\esmserver\inkpixi\ONLINE_MARKETING\Salesforce Marketing Cloud\EXPORTS\from Exact Target', 'ExactTarget - Export Control Sheet.csv')
        os.startfile(ss)        
    
    def thread_error(self, txtError):
        #To show errors to user when something goes wrong...work in progress
        QMessageBox.critical(self, 'I got a bad feeling about this...', 'Something went wrong here... \n ' + txtError, QMessageBox.Ok)
    
class SendEmail(QThread):
        #QThread to send email after a file has been uploaded to Cross Country's FTP servers.
        emailSent = pyqtSignal()
        
        def __init__(self, rowCount = 0, lastOrder = None):
            super(QThread, self).__init__()
            self.rowCount = "{:,}".format(rowCount)
            self.lastOrder = datetime.strftime(lastOrder, '%m/%d/%Y')
            
        def run(self):
            #Set up addresses for email.
            toEmail, ccEmail, fromEmail = ['ESM_Implementation@crosscountrycomputer.com', 'ainkles@crosscountrycomputer.com'], 'cle@earthsunmoon.com', 'jwitmer@earthsunmoon.com'
            #toEmail, ccEmail, fromEmail = ['dhoy@aotees.com', 'hoy.davidj@gmail.com'], 'hoy.davidj@gmail.com', 'jwitmer@inkpixi.com'
            msg = MIMEMultipart('alternative')
            
            #set up and log into SMTP server.
            s = smtplib.SMTP('smtp.aotees.com', 1025)
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
        
    def __init__(self, sortedList):
        super(QThread, self).__init__()
        self.sortedList = sortedList
        
    def run(self):
        
        #Set up filename and path to create a file for uploading.
        today = time.strftime('%m%d%Y')
        upload_file_name = 'PIXI_DMCID_subs_' + today + '.csv'
        upload = os.path.join(r'upload', upload_file_name)  
        remote_path = posixpath.join('ccftpEsM', upload_file_name)
                
        #Write to the file from the list create in ProcessCrossCountry()
        try:
            with open(upload, 'w', newline='') as up:
                wr = csv.writer(up)
                wr.writerow(['Email Address', 'DMCID', 'Last_order_date_HH', 'Last_Order_DB_HH'])
                for li in self.sortedList:
                    wr.writerow([li[0], li[1], li[2], li[3]])
                up.close()
            
            #os.startfile(upload)
            #Gets the total size of the file for tracking upload.
            self.totalSize = os.path.getsize(upload)
            
            #Upload the file, with a handler to track the progress.
            cnopts = pysftp.CnOpts()
            cnopts.hostkeys = None
            with pysftp.Connection('sftp.crosscountrycomputer.com', username='ccftpEsm', password='STia6#Rx', cnopts=cnopts) as sftp:
                #sftp.put(upload, self.handle)
                sftp.put(upload, remote_path, self.handle)
            #move to server for storage
            serverLocation = os.path.join(r'\\esmserver\inkpixi\ONLINE_MARKETING\Salesforce Marketing Cloud\EXPORTS\from Exact Target\PIXI_DMCID_Subs', upload_file_name)

            shutil.move(upload, serverLocation)              
            
            #Tell caller thread finished successfully.
            self.uploadFinished.emit()
            #change FileExistsError to BaseExcepction before production.
        #except FileExistsError as e:
        except BaseException as e:
            #Tell the caller something went wrong, usually a permissions error because someone has the sheet open.
            self.error.emit(str(e))
            
    def handle(self, transferred, to_transfer):
        #Use to track the progress of the upload in percent.
        percentComplete = round((transferred / to_transfer) * 100)
 
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
#         with open(self.cross_file, 'r') as f:
#             reader = csv.reader(f)
#             #get's header row in csv file
#             headers = next(reader)
#             f.close()
        headers = ['Email Address', 'DMCID', 'Last_order_date_HH', 'Last_Order_DB_HH']
        
        #Open file and dump into list for sorting, probably a better way to do this.
        with open(self.cross_file, 'r') as f:
            reader = csv.reader(f)
            #Get rid of headers.
            next(reader)
            lsDt = []
            for dt in reader:
                if dt[4]:
                    #Format strings back into dates for sorting.
                    #lsDt.append(datetime.strptime(dt[2], '%m/%d/%Y %I:%M %p'))
                    lsDt.append(datetime.strptime(dt[4], '%m/%d/%Y'))
            f.close()
        maxDt = max(lsDt)
        
        #Set up file to write tracking information for marketing.
        fa = os.path.join(r'\\esmserver\inkpixi\ONLINE_MARKETING\Salesforce Marketing Cloud\EXPORTS\from Exact Target', 'ExactTarget - Export Control Sheet.csv')
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
                dt = self.get_sort_date(d[4])
                lsCsv.append([d[0], d[3], dt, 'IP'])
            
            #Sort the list by "Last Order Date" descending.
            #sortedList = sorted(lsCsv, key= lambda row: datetime.strptime(row[2], '%m/%d/%Y %I:%M %p'), reverse = True)
            sortedList = sorted(lsCsv, key= lambda row: datetime.strptime(row[2], '%m/%d/%Y'), reverse = True)
            sort.close()        
        
        
        #Send back the list.
        self.process_results.emit(sortedList, row_count, maxDt)
        
    def get_sort_date(self, getDate):
        #If there is no date, assign a minimum date.
        #minDate = '1/1/1900 12:00 AM'
        minDate = '1/1/1900'
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
        elif self.import_type == 'update_dmcid':
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
        
        #Not sure why this one starts at one, but too scared to change it right now.
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
        
class GetTopFifty(QThread):
    top_fifty = pyqtSignal(list)
    
    def __init__(self, sku, from_dt, to_dt):
        super(GetTopFifty, self).__init__()
        self.sku, self.from_dt, self.to_dt = sku, from_dt, to_dt
        
    def run(self):
        con = ConnectData()
        cur = con.connect_inkpixi() 
        params = (self.from_dt, self.to_dt, self.sku)
        cur.execute("{CALL dbo.uspGetTopFiftyByGarmType (?, ?, ?)}", params)
        ds = cur.fetchall()

        self.top_fifty.emit(ds)
        
class ConnectData(object):
    #Only exists to connect to database.
    def connect_dmc(self):
        conn = pyodbc.connect('DRIVER={SQL Server}; SERVER=SQLSERVER; DATABASE=DMC; Trusted_Connection=yes')
        db = conn.cursor()
        
        return db    
    
    def connect_online_marketing(self):
        conn = pyodbc.connect('DRIVER={SQL Server}; SERVER=SQLRPTSERVER\SQLREPORTS; DATABASE=OnlineMarketing; Trusted_Connection=yes')
        db = conn.cursor()
        
        return db     
    
    def connect_import_export(self):
        conn = pyodbc.connect('DRIVER={SQL Server}; SERVER=SQLSERVER; DATABASE=ImportExport; Trusted_Connection=yes')
        db = conn.cursor()
        
        return db
    
    def connect_inkpixi(self):
        conn = pyodbc.connect('DRIVER={SQL Server}; SERVER=SQLSERVER; DATABASE=InkPixi; Trusted_Connection=yes')
        db = conn.cursor()
        
        return db                            
    
class MatchCampaigns(QThread):
    send_results = pyqtSignal(list, int)
                           
    def __init__(self):
        super(QThread, self).__init__()
        
    def run(self):
        #Set up connection to database.
        con = ConnectData()
        cur = con.connect_dmc()
        #Grab a list of all email address with the selected campaign id.
        with closing(cur) as db:
            db.execute("EXEC dbo.get_dmc_campaign_from_email")
            ds = db.fetchall()
        
        row_count = 0
        ls_matches = []
        for i in ds:
            ls_matches.append([i[0], i[1], i[2]])
            row_count += 1
        
        #Tell caller I'm finished.
        self.send_results.emit(ls_matches, row_count) 
        
class MatchDMCID(QThread):
    #Returns list of people who aren't matched to a campaign.
    #send_export = pyqtSignal(list, int)
    match_finished = pyqtSignal()
    def __init__(self):
        super(QThread, self).__init__()
        
    def run(self):
        con = ConnectData()
        cur = con.connect_dmc()
        with closing(cur) as db:
            db.execute("EXEC dbo.UpdateDMCIDs")
            #ds = db.fetchall()
        self.match_finished.emit()
#         row_count = 0
#         print(ds)
#         ls_non_match= []
#         for i in ds:
#             ls_non_match.append([i[0], i[1]])
#             row_count += 1
#             
#         self.send_export.emit(ls_non_match, row_count) 
                
if __name__ == "__main__":
    myappid = 'Exact Target Email Compare'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid) 
    
    app = QApplication(sys.argv)
    app.setStyle("plastique")
    fw = FTPWindow()
    fw.show()
    
    sys.exit(app.exec_())