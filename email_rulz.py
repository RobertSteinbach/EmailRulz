# https://pypi.org/project/imap-tools/    IMAP Tools webpage

from imap_tools import MailBox, AND, OR  # Needs PIP INSTALL
import imap_tools
import imaplib
import email.message
import sqlite3
import time  # for sleep
import datetime
import re               # maybe
import os               # to get environment variables

#some even moreto test commit


# Global scope

# Get secrets from environment variables
#   In Pycharm, select the "Edit Configuration" menu item in the project drop-down box in top menu bar
imapserver = os.environ.get("IMAP_SERVER")
userid = os.environ.get("IMAP_LOGIN")
pd = os.environ.get("IMAP_PWD")
myemailaddress = os.environ.get("EMAIL_ADDRESS")
#print(imapserver, userid, pd, myemailaddress)          #print out credentials to make sure I got them

# Everything else
autofilefolder = "INBOX.autofile"       #Everything occurs under this subfolder
runlog = []  # Build a list of events
SubFolders = []  # List of subfolders below the autofile folder.
SubfolderPrefix = autofilefolder + "."
status = ""
ActionFields = ("from", "subject", "body")
rundt = datetime.datetime.now()
extractrulesflag = False
sleeptime = 300  # 300 seconds = 5 minutes between iterations

def cleanup():

    #Clean out old log files.  Criteria:
    # 1.  From myself
    # 2.  Incremental logs
    # 3.  Older than 2 days (today and yesterday)
    # 4.  Unflagged (no errors)

    status = "Cleanup: delete old incremental log emails..."
    runlog.append(status)

    date_criteria = (datetime.datetime.now() - datetime.timedelta(days=1)).date()
    #print("date_criteria-", date_criteria)

    try:
        mailbox.folder.set(autofilefolder)      #point mailbox to autofile folder
        numrecs = mailbox.delete(mailbox.fetch(AND(from_=[myemailaddress], subject=["Rulz Log - Incremental @"],
                                   date_lt=date_criteria, flagged=False), mark_seen=False))
        runlog.append(str(str(numrecs).count('UID')) + " log messages removed.")
    except Exception as e:
        runlog.append("!!!FAILED to cleanup old log files from autofile folder.")
        runlog.append(str(e))
        return


    return  #end of cleanup



def looper():
    a = 1
    while a > 0:
        a += 1              # a will be the time period between iterations (e.g. 5 minutes)

        # extract the rules on the first iteration and about every 3 days
        if a == 2 or (a % 1000 == 0):
            extract_rulz()

        # clean up the autofile folder on the first iteration and about every 8 hours
        if (a == 2) or (a % 100 == 0):
            cleanup()

        # refresh the list of subfolders; strip off the prefix
        SubFolders.clear()
        for folder_info in mailbox.folder.list(SubfolderPrefix):
            SubFolders.append(folder_info['name'].replace(SubfolderPrefix, ''))

        # check to see if any rules need to be changed
        change_rulz()
        # return

        # Process rules on INBOX
        process_rulz()

        # Dump the Event log to an email
        emailflag = ""  # Assume no flags
        if str(runlog).find("!!!") > -1:
            runlog.append("Errors were found.  Adding FLAG to the log.")
            # emailflag = imap_tools.MailMessageFlags.FLAGGED    # Flag it there was an error (marked by !!!)
            emailflag = "\FLAGGED"
        new_message = email.message.Message()
        new_message["To"] = myemailaddress
        new_message["From"] = myemailaddress
        new_message["Subject"] = "Rulz Log - Incremental @ " + str(rundt)
        new_message.set_payload('\n'.join(map(str, runlog)))
        mailbox2.append(autofilefolder, emailflag, imaplib.Time2Internaldate(time.time()),
                        str(new_message).encode('utf-8'))

        # Dump out the event log to the screen
        print("************* EVENT LOG **************")
        for event in runlog:
            print(event)

        # Clear the run log
        runlog.clear()

        # Sleep for a while
        print("Sleeping for " + str(sleeptime) + " seconds...")
        time.sleep(sleeptime)  # 600 = 10 minutes
    return  # Bottom of endless loop


def extract_rulz():
    # status
    runlog.append("Extracting rules to email...")

    # extract the rules and send back
    sql = "select field, criteria, tofolder from rulz " \
          "order by field, criteria, tofolder;"
    cursor.execute(sql)
    rulz = cursor.fetchall()
    # for row in rulz:
    #    print(' '.join(row))
    # print('\n'.join(' '.join(row) for row in rulz))    #Dump the recordset (SAVE THIS LINE!)

    # Create a new mail with all the rulz
    new_message = email.message.Message()
    new_message["To"] = myemailaddress
    new_message["From"] = myemailaddress
    new_message["Subject"] = "rulz extract"
    new_message.set_payload("rulz:\n" + '\n'.join(' '.join(row) for row in rulz))
    mailbox2.append(autofilefolder, '', imaplib.Time2Internaldate(time.time()),
                    str(new_message).encode('utf-8'))
    return  # end of extract_rulz


def change_rulz():
    rundt = datetime.datetime.now()
    runlog.append(str(rundt) + " - checking for rule changes...")
    extractrulesflag = False

    msgs2move = []  # build a list of messages from myself to move.  Can't move them while in the loop of
    # messages b/c it will invalidate the recordset and next loop will fail

    # Get any mail sent from myself
    try:
        mailbox.folder.set('INBOX')
        mymsgs = mailbox.fetch(AND(from_=[myemailaddress]))
    except Exception as e:
        status = "!!! ERROR fetching messages from myself.  Error = " + str(e)
        runlog.append(status)
        return

    for msg in mymsgs:

        # Get the unique id
        uid = msg.uid

        if msg.subject.lower() == "rulz":  ### DUMP RULES TO EMAIL  ####

            extractrulesflag = True

            # Move the processed msg to the autofile folder to mark as processed
            msgs2move.append(uid)
            # mailbox.move(msg.uid, autofilefolder)

            continue  # onto next message

        # REPLACE all rules from email
        if 'rulz extract' in msg.subject:
            # replace all the rules being sent back

            # status
            runlog.append("Replacing rules...")

            rulzpos = msg.text.find('rulz:')
            if rulzpos < 0:
                status = "!!! ERROR - Could not find the 'rulz:' keyword.   Ignoring file."
                runlog.append(status)

                mailbox.flag(msg.uid, imap_tools.MailMessageFlags.FLAGGED, True)
                msgs2move.append(uid)
                # mailbox.move(msg.uid, autofilefolder)  # Move message as processed
                continue  # onto the next msg

            # The rulz_new table should NOT exist, but attempt to rename just in case
            sql = "ALTER TABLE rulz_new RENAME TO rulz_new_" + datetime.datetime.now().strftime("%b%d%Y%H%M%S") + ";"
            try:
                cursor.execute(sql)  # drop the temp table
                dbcon.commit()
            except:
                # don't care if this fails 
                status = "Error archiving old rulz_new table.   This is normal."
                # runlog.append(status)

            # Create a temp table named rulz_new
            sql = "SELECT sql FROM sqlite_master WHERE name = 'rulz';"
            try:
                cursor.execute(sql)  # Get the CREATE statement for the rulz table
                sql = cursor.fetchall()[0][0].replace('rulz', 'rulz_new')  # make a copy
                cursor.execute(sql)  # create new table
                dbcon.commit()  # do it
            except Exception as e:
                status = "!!! Error - could not find schema for 'rulz' table."
                runlog.append(status)

                mailbox.flag(msg.uid, imap_tools.MailMessageFlags.FLAGGED, True)
                msgs2move.append(uid)
                # mailbox.move(msg.uid, autofilefolder)  # Move message as processed
                continue  # onto the next msg

            # Build a list of tuples
            temprulz = msg.text[rulzpos + 7:].strip()  # Substring everything past the rulz: tag
            temprulz = temprulz.split('\r\n')  # Create a list from all the lines
            newrulz = []  # start with empty list
            for row in temprulz:  # each line now needs to put into a tuple
                # newrulz.append(tuple(str(row).strip().split(' ')))    # works, version #1

                # print(row)

                # https://stackoverflow.com/questions/2785755/how-to-split-but-ignore-separators-in-quoted-strings-in-python
                row_aslist = re.split(''' (?=(?:[^'"]|'[^']*'|"[^"]*")*$)''', row)  # I don't get it, but it works

                # parse it out into variables and evaluate them
                actionfield= str(row_aslist[0]).lower()
                if row_aslist[0] not in ActionFields:
                    status = "!!! ERROR parsing rule.  First word not recognized - " + actionfield
                    runlog.append(status)
                    runlog.append(row)
                    continue
                row_aslist[0] = actionfield           # force it to lowercase

                actioncriteria = str(row_aslist[1])     # add any validation rules here

                tofolder = str(row_aslist[2]).lower()
                if tofolder not in SubFolders:
                    status = "!!! ERROR parsing rule.  Target folder not recognized - " + tofolder
                    runlog.append(status)
                    runlog.append(row)
                    continue
                row_aslist[2] = tofolder            # force it to lower case

                # put the values in a tuple and then add it to the list
                newrulz.append(tuple(row_aslist))

            # newrulz=[('aaa','bbb','ccc'),('ddd','eee','fff')]    # this is the expected format
            sql = "INSERT INTO rulz_new (Field,Criteria,ToFolder) VALUES (?,?,?)"
            try:
                cursor.executemany(sql, newrulz)
                dbcon.commit()
            except Exception as e:
                status = "!!! ERROR inserting new data to rulz_new.  Error=" + str(e)
                runlog.append(status)
                runlog.append(sql)
                # status = 'New rules=' + str(newrulz)
                # runlog.append(status)
                for row in newrulz:
                    runlog.append(row)

                mailbox.flag(msg.uid, imap_tools.MailMessageFlags.FLAGGED, True)
                msgs2move.append(uid)
                # mailbox.move(msg.uid, autofilefolder)  # Move message as processed
                continue

            # Make a copy of the current Rulz_new table
            try:
                sql = "ALTER TABLE rulz RENAME TO rulz" + datetime.datetime.now().strftime("%b%d%Y%H%M%S") + ";"
                cursor.execute(sql)  # drop the temp table

                sql = "ALTER TABLE rulz_new RENAME TO rulz;"
                cursor.execute(sql)
                dbcon.commit()

            except Exception as e:
                status = "!!! ERROR attempting to archive/swap table 'rulz'. Error: " + str(e)
                runlog.append(status)

                mailbox.flag(msg.uid, imap_tools.MailMessageFlags.FLAGGED, True)
                msgs2move.append(uid)
                # mailbox.move(msg.uid, autofilefolder)  # Move message as processed
                continue

            # Move the processed msg to the autofile folder to mark as processed  - everything was good!
            msgs2move.append(uid)
            # mailbox.move(msg.uid, autofilefolder)

            # Extract the rules once more
            extractrulesflag = True

            continue  # onto next message
        # End REPLACE rulz from email

        #####################################################
        # CREATE ONE RULE FROM FORWARDED EMAIL
        #####################################################
        if (msg.subject.find('FW:') > -1) or (msg.subject.find('Fwd:') > -1):

            body = msg.text[:msg.text.find('\r')]                # get the first line of the body
            #print("Body=", body)

            # https://stackoverflow.com/questions/2785755/how-to-split-but-ignore-separators-in-quoted-strings-in-python
            body = re.split(''' (?=(?:[^'"]|'[^']*'|"[^"]*")*$)''', body)      #I don't get it, but it works

            # parse it out into variables
            actionfield = str(body[0]).lower().strip()
            actioncriteria = str(body[1])
            tofolder = str(body[2]).lower().strip()
            # print(actionfield)
            # print(actioncriteria)
            # print(tofolder)


            # If the actioncriteria was a hyperlink, then fix that
            if tofolder.find("mailto:") > -1:
                runlog.append("Criteria found to be a hyperlink.  Skipping over 'mailto:' tag.")
                # tofolder = msg.text.split()[3].lower()
                tofolder = tofolder.replace("mailto:", "")               #remove the mailto: tag

            status = "FW email found. ActionField='" + actionfield + "'. ActionCriteria='" + actioncriteria \
                     + "'. ToFolder='" + tofolder + "'."
            runlog.append(status)

            # make sure the first word is a valid action field (from, subject,...)
            if actionfield not in ActionFields:
                status = "WARNING - Did not find the first word '" + actionfield + "' to be a valid action field. " \
                        "Email was ignored. List of possible action fields are: " + str(
                    ActionFields)
                runlog.append(status)

                # mailbox.flag(uid, imap_tools.MailMessageFlags.FLAGGED, True)
                # mailbox.move(uid, autofilefolder)  # Move message as processed
                continue  # onto next message

            # make sure the tofolder is in the list of subfolders
            if tofolder not in SubFolders:
                # print(msg.text)
                status = "!!! ERROR - Did not find autofile folder '" + tofolder + ". Email was ignored. " \
                        "List of possible folders are: " + str(
                    SubFolders)
                runlog.append(status)

                mailbox.flag(uid, imap_tools.MailMessageFlags.FLAGGED, True)
                msgs2move.append(uid)
                # mailbox.move(uid, autofilefolder)  # Move message as processed
                continue  # onto next message

            # Create the rule in the database
            sql = "INSERT INTO Rulz (Field,Criteria,ToFolder) VALUES ('" + actionfield + "','" + actioncriteria + "'," \
                "'" + tofolder + "');"

            try:
                cursor.execute(sql)
                dbcon.commit()
            except Exception as e:
                status = "!!! ERROR - Could not insert new rule.  SQL='" + sql + \
                         "Error: " + str(e)
                runlog.append(status)

                mailbox.flag(uid, imap_tools.MailMessageFlags.FLAGGED, True)
                msgs2move.append(uid)
                # mailbox.move(uid, autofilefolder)  # Move message as processed

            # Move the msg to the autofile folder to mark as processed
            msgs2move.append(uid)
            # mailbox.move(uid, autofilefolder)

            # Give good status news
            runlog.append("Rule added! ID=" + str(cursor.lastrowid) + ". Action Field ='" + actionfield \
                          + "'. Criteria='" + actioncriteria + "'.  ToFolder='" + tofolder + "'.")

            # Extract the rules once more
            extractrulesflag = True

            continue  # to the next message
    # for each message sent from myself

    # Move all the processed messages from myself
    mailbox.move(msgs2move, autofilefolder)

    # If something changed, extract the rules again
    if extractrulesflag == True:
        extract_rulz()

    return  # end of change_rulz()


def process_rulz():
    # make a timestamp for the run
    rundt = datetime.datetime.now()
    runlog.append(str(rundt) + " - processing rules...")

    # Get the list of "to folders" from database.  Will move emails in bulk
    sql = "Select distinct ToFolder, Field from Rulz;"
    try:
        cursor.execute(sql)
        ToFolders = cursor.fetchall()
    except Exception as e:
        runlog.append("!!! ERROR - Could not get list of ToFolders.  Error=" + str(e))
        return

    for row in ToFolders:  # For every To Folder/Keyword....
        ToFolder = row[0]
        ToFolderVerbose = SubfolderPrefix + ToFolder
        actionField = row[1]

        sql = "select criteria from rulz where tofolder='" + ToFolder + "' AND field='" + actionField + "';"
        cursor.execute(sql)
        CriteriaSet = cursor.fetchall()
        Criteria = []
        for row2 in CriteriaSet:
            Criteria.append(row2[0].replace('"', ''))        # drop any double quotes in the criteria

        # Pull the emails that have the criteria
        # for msg in mailbox.fetch(OR(from_=['little', 'newegg', 'drafthouse.com']), mark_seen=False):
        status = "Processing '" + actionField + "' rules for '" + ToFolder + "' folder..."  # Criteria=" + str(Criteria)
        runlog.append(status)
        numrecs = 0
        try:
            mailbox.folder.set('INBOX')
            if actionField.lower() == "from":
                numrecs = mailbox.move(mailbox.fetch(OR(from_=Criteria), mark_seen=False), ToFolderVerbose)
            if actionField.lower() == "subject":
                numrecs = mailbox.move(mailbox.fetch(OR(subject=Criteria), mark_seen=False), ToFolderVerbose)
            if actionField.lower() == "body":
                numrecs = mailbox.move(mailbox.fetch(OR(body=Criteria), mark_seen=False), ToFolderVerbose)

            runlog.append("..." + str(str(numrecs).count('UID')) + " messages moved to folder '" + ToFolder + "'.")
        except Exception as e:
            runlog.append("!!!FAILED rule for folder '" + ToFolder + "'. Criteria=" + str(Criteria))
            runlog.append(str(e))
            continue

        # end for each keyword (from, subject, body...)
    # end for each ToFolder
    runlog.append("Rules processing completed.")
    return  # end of process_rulz()


def tester():
    print("test completed")
    quit()
    return  # end of tester()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    # tester()

    'Give feedback on start'
    status = 'Email_Rulz program started at: ' + datetime.datetime.now().strftime("%Y-%M-%d %H:%M:%S")
    print(status)
    runlog.append(status)

    # Connect to mailbox
    mailbox = MailBox(imapserver)
    try:
        mailbox.login(userid, pd)
        runlog.append("Mailbox connected")
    except Exception as e:
        runlog.append("Mailbox NOT connected.  Error=" + str(e))

    # Open a 2nd connection using the imaplib library.
    # This will allow me to create new email (imap-tools doesn't support)
    # TODO:  Not eloquent to login to the imap server twice

    mailbox2 = imaplib.IMAP4_SSL(imapserver)
    try:
        mailbox2.login(userid, pd)
        runlog.append("Mailbox2 connected")
    except Exception as e:
        runlog.append("!!! ERROR Mailbox2 NOT connected.  Error=" + str(e))

    # Connect to database
    try:
        dbcon = sqlite3.connect('./rulz/rulz.db')
        cursor = dbcon.cursor()
        runlog.append("Database connected")
    except Exception as e:
        runlog.append("!!! ERROR Database NOT connected.  Error=" + str(e))

    # Check for errors, abend if there were any errors to this point
    if "!!!" in runlog:
        print("Fatal error detected! Abend. ")
        for event in runlog:
            print(event)
        quit()

    # Call the endless loop
    looper()

    # Should never get here
    print('!!!!END OF PROGRAM!!!!')


def ParkingLot():
    """
        #   print("****FOLDERS:")
        #   for folder_info in mailbox.folder.list():
        #       print("folder_info", folder_info)

        for msg in mailbox.fetch(OR(from_=['little', 'newegg', 'drafthouse.com']), mark_seen=False):
            print("***MESSAGE:")
            print("UID-", msg.uid)
            print("Subject-", msg.subject)
            print("From-", msg.from_)
            # print("To-", msg.to)
            # print("CC-", msg.cc)
            print("ReplyTo-", msg.reply_to)
            # print("Flags-", msg.flags)
            # print("Headers-", msg.headers)
            # print("Text-", msg.text)

            print("** PROCESS RULES:")
            if msg.from_.find("littlecaesars") > -1:
                print("Found little caesers in msg ", msg.uid)
                try:
                    # mailbox.move(msg.uid, 'INBOX.Autofile.Junk')
                    print("message moved - ", msg.uid)
                except Exception as e:
                    print('Error moving message id ', msg.uid, '. Error-', e)
    
        
    #Create email via IMAP 
    https://stackoverflow.com/questions/3769701/how-to-create-an-email-and-send-it-to-specific-mailbox-with-imaplib  
        import imaplib
        connection = imaplib.IMAP4_SSL(HOSTNAME)
        connection.login(USERNAME, PASSWORD)
        new_message = email.message.Message()
        new_message["From"] = "hello@itsme.com"
        new_message["Subject"] = "My new mail."
        new_message.set_payload("This is my message.")
    
    
    #Python tips - How to easily convert a list to a string for display
    https://www.decalage.info/en/python/print_list
    
    #For statement on one line (list comprehensions)
    https://stackoverflow.com/questions/1545050/python-one-line-for-expression  
    
    
    """
