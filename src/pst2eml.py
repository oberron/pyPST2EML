"""
CONVERTS OUTLOOK PST FILE into directory of .eml files which can be indexed with desktop search tools

pst2eml.py (this file): changes recursively the .eml file name and properties (creation date, from, ...) in a given folder
    to match the properties of the original email (.eml, .ics) to allow desktop search filtering 

make_eml_search_friendly is the entry point
-> calls eml_get_parameters to extract from the .eml file the :
        send_to, sent_from, subject, sent_date parameters
-> then calls rename_eml (which in turns might call incrementalfilename to avoid filename duplicates)
-> then calls change_creation_date() and set_file_attributes() to allow desktop search like X1 or Lucene
    to index the files properly


Example #1 - convert a given pst file:
--------------------------------------
python .\src\pst2eml.py --folder=C:\outlook\archives --pst=Y --filename="2020_Q1_Q3.pst"

Example #2 - make the output of libpst searchable:
--------------------------------------------------
python .\src\pst2eml.py --folder=C:\outlook\archives\eml\2020_Q1_Q3 -v=4"

"""

""" PYTHON STANDARD LIBRARY """
import argparse
import email
from email.header import decode_header
import logging
import os
from os.path import abspath,  basename, dirname, exists, join, pardir, splitext
from pathlib import Path
from quopri import decodestring
from subprocess import call
from time import mktime, sleep, strptime

""" REQUIRED INSTALLS """
from dateutil.parser import parse as dateutil_parse

try:
    import win32file
    import win32con
    from win32com import storagecon
    import pywintypes
    import pythoncom
except:
    print("you might need to install pywin32")
    print(">python -m pip install pywin32")
    raise

folder = abspath(join(dirname(__file__), "test"))
stop_error = False

def scan_email_property(msg_strings,prop_vals):
    result = ""
    for line in msg_strings:
        for prop_val in prop_vals:
            if line.find(prop_val) >= 0:
                result = ":".join(line.split(":")[1:])
                return result
    return result

def scan_email_receive_header(msg_strings):
    dows = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    result = ""
    second_line = msg_strings[1]
    for dow in dows:
        if second_line.find(dow)>=0:
            result = dow+second_line.split(dow)[1]
            return result
    return result

def eml_get_parameters(eml_fp,debug=False):
    """ Gets the Params from the email stored as text file on HDD
    Parameters:
    -----------
    eml_fp: str
        full path to the .eml file
    Returns: 
    --------
    send_to: str
    from_sender: str
    title: str
    sent_date : str
    """
    SendTo = ""
    try:
        if eml_fp.find(".ics")>=0:
            msg = email.message_from_file(open(eml_fp,encoding="utf-8"))
            msg_file = open(eml_fp,'r',encoding="utf-8")
            msg_strings = msg_file.readlines()
            msg_file.close()
        else:
            msg = email.message_from_file(open(eml_fp,encoding="latin-1"))
            msg_file = open(eml_fp,'r',encoding="latin-1")
            msg_strings = msg_file.readlines()
            msg_file.close()
    except:
        log_msg = "failed loading email : %s"%(eml_fp)
        logging.error(log_msg)
        raise

    try:
        SendTo = msg.get("TO")
    except:	
        SendTo = ""

    if len(str(SendTo))==0:
        # msg_file = open(emlfile,'r')
        # msg_strings = msg_file.readlines()
        # msg_file.close()
        for line in msg_strings:
            if line.find("To:") >= 0:
                SendTo = line.split(":")[1]
    if SendTo == "" or SendTo is None:
        logmsg = "Cannot get TO: field from :%s"%(eml_fp)
        logging.info(logmsg) #we do not raise an exception as the to: field is optional for the windows desktop search
        SendTo = ""

    ######## FROM

    try:
        From = msg.get("FROM")
    except:		
        for line in msg_strings:
            if line.find("From:") >= 0:
                From = line.split(":")[1]
        if From == "":
            log_msg = "Cannot get FROM: field from :%s"%(eml_fp)
            logging.error(log_msg)
            raise Exception("HeaderError",log_msg)
    ######## SUBJECT 
    """ Gets the Subject topic """
    Subject = None
    try:
        Subject = msg.get("SUBJECT")
        log_msg = "read subject from message: %s"%(Subject)
        logging.debug(log_msg)
    except:
        log_msg = "failed to get MSG for : %s"%(eml_fp)
        logging.error(log_msg)

    if Subject == "":
        log_msg = "failed to get not empty string as Subject for : %s"%(eml_fp)
        logging.debug(log_msg)

        for line in msg_strings:
                if line.find("Subject:") >= 0:
                    Subject = line.split(":")[1]
        if len(Subject)==2:
            if Subject.replace("\n","").replace("\t","").replace(" ","")=="":
                Subject = ""
    if Subject == "":
        log_msg = "failed to get not empty string as Subject, now looking for filename for : %s"%(eml_fp)
        logging.debug(log_msg)

        #enter here if nowhere in the file there is a line "Subject"
        # we look for a file attachment to name the email 
        for line in msg_strings:
            if line.find("filename=") >= 0:
                Subject = line.split("=")[1].split('"')[1]

    if Subject == "" or Subject is None and eml_fp.find(".ics")>=0:
        for line in msg_strings:
            if line.find("SUMMARY") >= 0:
                Subject = ":".join(line.split(":")[1:]).replace("\n","")

    if Subject == "" or Subject is None:
        LogMessage = "Cannot get SUBJECT: field from :%s"%(eml_fp)
        logging.warning(LogMessage)
        Subject="No Subject"

    default_charset = 'UTF-8'
    default_charset = 'latin-1'
    subject_lines = decode_header(Subject)
    #https://docs.python.org/3/library/stdtypes.html#str
    #class str(object=b'', encoding='utf-8', errors='strict')
    try:
        subject = ""
        for subject_string, subject_encoding in subject_lines:
            if not subject_encoding:
                #when encoding is None, just append the fragment
                if type(subject_string)==bytes:
                    subject+=str(subject_string,default_charset)
                else:
                    subject+= subject_string
            else:
                #TL;DR: japanese encoding in outlook is a mess
                if subject_encoding.upper() in ["ISO-2022-JP","ISO-2022-JP-1", "ISO-2022-JP-2",\
                    "ISO-2022-JP-3", "ISO-2022-JP-2004", "CP932","WINDOWS-31J",
                    "ShiftJIS","CP942","SJIS","UCS2","SHIFT_JISX0213","SHIFT_JISX0208"]:
                    subject_encoding = "CP932"
                res = subject_string.decode(subject_encoding,"ignore")
                subject+= res
    except:
        log_msg = "failed to extract subject for : %s"%(eml_fp)
        logging.error(log_msg)
        raise

    try:
        subject = subject.replace("\n"," ").replace("\t"," ").replace("FW: ","").replace("RE: ","").replace(":"," ").replace("/","-").replace("\\","-").replace("*","-").replace("?","-").replace("\"","'").replace("<","-").replace(">","-").replace("|","-").replace("\n"," ")
        #frequent encoding errors
        subject = subject.replace("Ã©","é")
    except TypeError:
        log_msg = "Type Error for: %s"%(eml_fp)
        logging.error(log_msg)
        raise

    #remove characters from 1 to 31
    #source https://docs.microsoft.com/en-us/windows/win32/fileio/naming-a-file
    subject = "".join([c for c in subject if ord(c)>31])
    if len(subject)==0:
        subject="NoSubject"


    ######## SENT DATE
    """ Gets the Send_Date """
    SentDate = None
    SentDate_property = scan_email_property(msg_strings,["Date:","Sent: ","DTSTART"])
    SentDate_header =  scan_email_receive_header(msg_strings)
    if debug:
        print(f"\t\tSentDate_property: {SentDate_property}")
        print(f"\t\tSentDate_header: {SentDate_property}")
    try:
        #most appropriate is to get the sent date from the email
        SentDate  = msg.get("Date")
    except:
        SentDate=None
    if SentDate is None:
        if len(SentDate_property)>len(SentDate_header):
            SentDate = SentDate_property
        else:
            SentDate = SentDate_header
    if debug:
        print(f"\t\tSentDate: {SentDate}")
    if SentDate  == "" or SentDate is None:
        logmsg = "Cannot get SENT DATE: field from :%s"%(eml_fp)
        logging.error(logmsg)
        raise Exception("HeaderError",logmsg)

    SentDate = SentDate.replace("\n","")
    SentDate = SentDate.replace("\r","")
    SentDate = SentDate.lstrip()
    SentDate = SentDate.rstrip()

    #handle string formatting not handled by dateutil_parse
    for tz in [" W. Europe Standard Time"," (GMT)", ' "GMT"']:
        if SentDate.find(tz)>=0:
            SentDate = SentDate.replace(tz,"")
    #handle DOW not handled by dateutil_parse
    weird_dow = {"Wen":"Wed"}
    for dow in weird_dow:
        if SentDate.find(dow)>=0:
            SentDate=SentDate.replace(dow,weird_dow[dow])
    if debug:
        print("SentDate before dateutil_parse",SentDate)
    try:
        emailtime = dateutil_parse(SentDate)
        if debug:
            print("\t\temailtime parsed by dateutil_parse",emailtime)
    except:
        logmsg = "Cannot get SENT DATE: field from :%s"%(eml_fp)
        logging.error(logmsg)
        raise Exception("HeaderError",msg)


    #sent_to,from_sender, title, sent_date
    return SendTo, From, subject, SentDate

def incrementalfilename(path,fn):
    """ create a new filename to avoid duplicates
    Parameters:
    -----------
    Return:
    -------
    new_fn: str
        new file name (with extension)
    """
    #fnb: filename
    fnb='.'.join(fn.split(".")[:-1])
    #fne: file extension
    fne = fn.split(".")[-1]
    #idea is we keep incrementing until a first value for which 
    #no existing file with same name in folder can be found
    final =0
    for i in range(1,1024):
        if final==0:		
            #print path+fnb
            new_fn = fnb+'['+str(i)+'].'+fne
            new_fp = abspath(join(path, new_fn))
            if not(os.path.isfile(new_fp)):
                final = i
                break
    #new_fn is filename + [1] + . eml
    new_fn = fnb+'['+str(final)+'].'+fne
    return new_fn

def rename_eml(eml_fp,subject,ignore_ext = True, debug=False):
    """ rename the file on the HDD
    Parameters:
    -----------
    eml_fp: str
        current full path to file
    subject: str
        subject of email and future file name
    Returns:
    --------
    new_eml_fp: str
        full path (folder and filename with extension)
    """
    folder_path = dirname(eml_fp)
    fp_ext = Path(eml_fp).suffix
    if ignore_ext:
        fp_ext = ".eml"

    new_eml_fp = abspath(join(folder_path,subject+fp_ext))
    log_msg = f"315 subject: -{subject}-, extension -{fp_ext}-,\n\t\t fp -{new_eml_fp}-"
    if debug:
        print(log_msg)
    logging.debug(log_msg)
    #windows seems to still have issues renaming when full path length >260
    #https://docs.microsoft.com/en-us/windows/win32/fileio/naming-a-file#maximum-path-length-limitation
    if len(new_eml_fp)>260:
        subject = subject[:256-len(folder_path)]
        new_eml_fp = abspath(join(folder_path,subject+fp_ext))
    if os.path.isfile(new_eml_fp):
        #check if the file name already exists and create a new one
        new_filename = incrementalfilename(folder_path,subject+fp_ext)
        new_eml_fp = abspath(join(folder_path,new_filename))
    if debug:
        print(f"new_eml_fp {new_eml_fp}")
    try:
        os.rename(eml_fp,new_eml_fp)
    except:
        LogMessage = "cannot rename file: %s - %s"%(eml_fp,new_eml_fp)
        logging.error(LogMessage)
        raise
    return new_eml_fp

def change_creation_date(eml_fp,sent_date):
    """ change file creation date to match sent_date
    Parameters:
    -----------
    eml_fp: str
        full path to the file
    sent_date: str
        date at which email was sent
    Returns:
    --------
    """
    emailtime = dateutil_parse(sent_date).timetuple()
    log_msg = "SENT ON: %s"%(str(emailtime))
    logging.debug(log_msg)

    atime = int(mktime(emailtime))
    times = (atime, atime)
    try:
        #new_fn_fp = abspath(join(dirname,new_filename)) #new filename full path
        #ChangeFileCreationTime(dirname+"\\"+new_filename,atime)
        #ChangeFileCreationTime(new_fn_fp,atime)
        wintime = pywintypes.Time(atime)
        try:
            #HERE WE TRY TO CHANGE the FILE CREATE DATE
            winfile = win32file.CreateFile(eml_fp, win32con.GENERIC_WRITE,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                None, win32con.OPEN_EXISTING,
                win32con.FILE_ATTRIBUTE_NORMAL, None)
            win32file.SetFileTime(winfile, wintime, None, None)
            winfile.close()
        except:
            log_msg = "179 error changing creation date for fname :%s"%(eml_fp)
            logging.info(log_msg)
            if stop_error:
                raise

        #NOW WE CHANGE THE LAST CHANGE TIME
        #setFileAttributes(eml_fp, from_sender, title,comments)
        os.utime(eml_fp, times)
    except:
        log_msg = "294 failed to change file flags for : %s"%(eml_fp)
        logging.info(log_msg)
        raise

def set_file_attributes(eml_fp, author, title,comments):
    """ change the windows file attributes (author, title, comments) to make easier for searching
    Parameters:
    -----------
    eml_fp: str
        full path to the email file on HDD
    Returns:
    --------
    None 
    """
    flags=storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE | storagecon.STGM_DIRECT
    pss=pythoncom.StgOpenStorageEx(eml_fp, flags, storagecon.STGFMT_FILE, 0 , pythoncom.IID_IPropertySetStorage,None)
    try:
        ps=pss.Create(pythoncom.FMTID_SummaryInformation,pythoncom.IID_IPropertyStorage,0,storagecon.STGM_READWRITE|storagecon.STGM_SHARE_EXCLUSIVE)
    except:
        try:
            ps=pss.Open(pythoncom.FMTID_SummaryInformation,storagecon.STGM_READWRITE|storagecon.STGM_SHARE_EXCLUSIVE)
        except:
            print("80 Failed \teml_fp: %s\n\tauthor:: %s\n\ttitle:: %s\n\tcomments::%s"%(eml_fp,author,title,comments))
            raise
    ps.WriteMultiple((storagecon.PIDSI_KEYWORDS,storagecon.PIDSI_COMMENTS,storagecon.PIDSI_AUTHOR,storagecon.PIDSI_TITLE),('keywords',comments,author,title))

    #add here wait loop to secure the filename change was effective
    while not os.path.isfile(eml_fp):
        sleep(0.01)
        
    ps=None
    pss=None

def is_eml(send_to, sent_from, subject, sent_date):
    if len(send_to)>3 and len(sent_from)>3 and len(subject)>=3 and len(sent_date)>8:
        return True
    else:
        return False

def str2bool(v):
    #credit https://stackoverflow.com/a/43357954/10567771
    if isinstance(v, bool):
       return v
    if v.lower() in ('yes', 'true', 't', 'y', '1'):
        return True
    elif v.lower() in ('no', 'false', 'f', 'n', '0'):
        return False
    else:
        raise argparse.ArgumentTypeError('Boolean value expected.')

def make_eml_search_friendly(eml_folder,ignore_ext = False): #,send_to, sent_from, subject, sent_date,OldArchives=False):
    """ rename and change files creation date and attributes  to reflect email content
	Parses files in emlpath and sets the file properties to the email properties
	if OldArchives ==True: considers that the files have been copied by the OS (file restore, back-up restore, ...) without the properties and will go through all .eml files
	Parameters:
    -----------
    eml_folder: str
        folder path to .eml file
    Returns:
    --------
    None
    """
    old_root = ""
    logging.debug(f"walking folder ignoring extension: {ignore_ext}")
    for root, directory, files in os.walk(eml_folder):
        for fn in files:
            if fn.find(".eml")>=0 or fn.find(".ics")>=0 or ignore_ext:
                if not root==old_root:
                    print(f"processing folder {root}")
                    old_root = root
                # if fn.find("utf-8")>=0 or fn.find("iso-8859")>=0 or fn.find("gb2312")>=0:
                log_msg = "processing:%s"%(fn)
                logging.info(log_msg)
                eml_fp = abspath(join(root,fn))
                send_to, sent_from, subject, sent_date = eml_get_parameters(eml_fp)
                log_msg = "\t\t\tsubject: %s"%(subject)
                logging.debug(log_msg)
                log_msg = "***************327***sent_date: %s"%(sent_date)
                logging.debug(log_msg)

                if (ignore_ext and is_eml(send_to, sent_from, subject, sent_date)) or not ignore_ext:
                    #if we ignore ext need to verify this is an email to give it an .eml extension
                    #in the case where other scripts have corrupted file names

                    try:
                        new_eml_fp = rename_eml(eml_fp,subject,ignore_ext)
                    except:
                        log_msg = "failed renaming file for: %s"%eml_fp
                        raise
                    try:
                        set_file_attributes(new_eml_fp, sent_from,subject,comments = "TO:"+send_to )
                    except:
                        log_msg = "failed updating file attrributes for: %s"%new_eml_fp
                        logging.error(log_msg)

                    try:
                        change_creation_date(new_eml_fp,sent_date)
                    except:
                        log_msg ="failed changing creation date: %s"%new_eml_fp
                        logging.error(log_msg)
                        if stop_error:
                            raise

def pst_2_eml(**kwargs):

    PST_folder = kwargs["folder"]
    PST_file   = kwargs["fn"]
    if "EML_folder" in kwargs:
        EML_folder = kwargs["EML_folder"]
    else:
        EML_folder = abspath(join(PST_folder,"eml"))
    #override_eml = kwargs["override_eml"]
    #old_archives = kwargs["old_archives"]

    pstlib_full_path = abspath(join(__file__,pardir,pardir,"bin","readpst.exe"))
    pst_full_path = abspath(join(PST_folder,PST_file))
    if exists(pst_full_path):
        if exists(pstlib_full_path):
            call([pstlib_full_path,"-e",pst_full_path,"-o",EML_folder])
        else:
            print("cannot find pstlib in given path:", pstlib_full_path)
            raise Exception(FileNotFoundError)
    else:
        print("pst file path not found:",pst_full_path)
if __name__=="__main__":
    parser = argparse.ArgumentParser(description='Processing emails for easier desktop searches')
    parser.add_argument('--pst', dest='PST_nEML',type=str2bool,
                    help='processing .PST (Y) or folder of .eml/.ics (N) (default: N)', default = "N")
    parser.add_argument('--folder', "-f", dest='folder',type=str , default = folder,
                    help="folder where processing will occur, if none then test folder used:")
    parser.add_argument('--filename',"-n", dest='fn',type=str,action="store", default = "c:/",
                    help='file to be processed (required for .PST processing, optional for .eml processing)')
    parser.add_argument('--embedded_libpst', dest='embedded_libpst',type=str, default = "Y",
                    help='use the package provided libpst (default: Y)')
    parser.add_argument("--test","-t",dest='test',type=str, default = "N",
                    help='will run over test files')
    parser.add_argument("--ignore","-i",dest='ignore',type=str2bool, default = "N",
                    help='ignore file extensions (if N only processes .eml and .ics)')
    parser.add_argument("--verbose","-v",dest='verbosity',type=int,default=2)

    args,  unknown = parser.parse_known_args()
    myargs = vars(args)

    if myargs["verbosity"]==4:
        logging.basicConfig(level=logging.DEBUG)
        logging.info("logging set to DEBUG")
        stop_error = True
    else:
        logging.basicConfig(level=logging.WARN)
        logging.info("logging set to WARN")

    if myargs["test"]=="Y":
        myargs["PST_nEML"]=False
        myargs["folder"]="D:\\gitw\\pyPST2EML\\test"


    if myargs["PST_nEML"]:
        pst_2_eml(**myargs)
    else:
        make_eml_search_friendly(myargs["folder"],myargs["ignore"])
