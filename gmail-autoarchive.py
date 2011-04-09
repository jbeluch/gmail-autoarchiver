#!/usr/bin/env python
'''
GMail AutoArchiver

This script will archive emails in your inbox over a certain age.

Usage:
    1. Set up filters in gmail matching the pattern 'autoarchive:\d+'
    where the \d+ is the age limit in days, e.g. 'autoarchive:3'.
    2. Run this script. If an email is in your inbox and is older than
    it's specified age limit, it gets archived.

Further Reading:
    GMail IMAP: http://code.google.com/apis/gmail/imap/
    IMAP Protocol: http://tools.ietf.org/search/rfc3501
    python imaplib: http://docs.python.org/library/imaplib.html

Hints about GMail:
    - GMail labels are = to IMAP folders.
    - To archive an email in gmail, you set the IMAP +FLAG \Deleted
'''

'''next features
- oauth instead of pw in file
- download all labels and use regex to match instead of autoarchive:*
'''

from datetime import datetime, tzinfo, timedelta
import imaplib
import email
import getpass

# Either hard code values here, or leave blank and you will be prompted
# upon script execution.
EMAIL_ADDRESS = ''
PASSWORD = ''

## First some helpful timezone stuff
# From http://docs.python.org/library/datetime.html
ZERO = timedelta(0)
class FixedOffset(tzinfo):
    """Fixed offset in minutes east from UTC."""

    def __init__(self, offset, name):
        self.__offset = timedelta(minutes = offset)
        self.__name = name

    def utcoffset(self, dt):
        return self.__offset

    def tzname(self, dt):
        return self.__name

    def dst(self, dt):
        return ZERO

utc = FixedOffset(0, 'UTC')

def connect(username, password):
    s = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    s.login(username, password)
    return s

def get_autoarchive_labels(s):
    '''Returns a list of tuples (str labelname, int age_in_days)'''
    _, list_of_labels = s.list(pattern='autoarchive:*')
    # Annoyingly if has no matches it returns a list with one element, None
    if list_of_labels[0] == None:
        return []

    # labels looks like: '(\\HasNoChildren) "/" "autoarchive:1"'
    # we want to extract the 'autoarchive:1' part
    ret = []
    #print list_of_labels
    for item in list_of_labels:
        label = item.split('"')[-2] # label = 'autoarchive:3'
        age = int(label.split(':', 1)[-1]) # age = 3
        ret.append((label, age))

    return ret

def get_message_ids(s, label):
    '''Takes an imap connection 's', and a label and returns a list
    of message ids for that label'''
    _, email_ids_string = s.search(None, 'X-GM-LABELS', label)
    # e.g. email_ids_string = ['2 3 4 8 11 14 15 17 18']
    email_ids = email_ids_string[0].split()
    return email_ids

def build_tz(tzstring):
    '''Takes a tzstring like '-0500 (EST)' or '-0500' and returns a
    datetime.tzinfo class''' 
    # messy :(

    # Using split(), the first item is always the offset, e.g. -0500
    parts = tzstring.split()
    offset = parts[0]
    
    # If there are 2 parts, the second part will be the tzname 
    name = 'UNKNOWN'
    if len(parts) > 1:
        name = parts[1].strip('()')

    # We need to convert offset into minutes and keep the +/- sign
    sign = offset[0]
    hours = int(offset[1:3])
    minutes = int(offset[3:5])
    offset_minutes = hours * 60 + minutes
    if sign == '-':
        offset_minutes *= -1
    tz = FixedOffset(offset_minutes, name)
    return tz

def archive_message(s, msg_id):
    ''' Simply set the deleted flag and msg will be archived in 
    gmail. '''
    s.store(msg_id, '+FLAGS', '"\\\\Deleted"')

def archive_old_messages(s, msg_ids, age):
    '''Takes a list of msg_ids and an age in days, and will archive
    any messages older than the age limit.'''
    now = datetime.utcnow().replace(tzinfo=utc)
    age_limit = timedelta(days=age)

    for msg_id in msg_ids:
        # Just get the date and subject message header, no need for entire msg
        _, headers = s.fetch(msg_id, '(body[header.fields (date subject)])')
        # Use email module, makes parsing headers easy :)
        e = email.message_from_string(headers[0][1])
        
        datestr = e.get('Date').strip()
        # e.g. datestr = 'Thu, 7 Apr 2011 08:34:04 -0400 (EDT)'
        subject = e.get('Subject').replace('\r\n', ' ')

        # %z doesn't work , must manually build timezone.
        # We want to remove the offset and sometimes present tzname from 
        # datestr so we can use strptime() 
        
        # someitmes datestr has '(tzname)' at the end and sometimes not, so we
        # split from left a set amount, guess i could use re
        parts = datestr.split(' ', 5)
        datetimestr = ' '.join(parts[:5])
        tzstring = parts[5]

        # Make dt object and apply the correct tzinfo
        dt_naive = datetime.strptime(datetimestr, r'%a, %d %b %Y %H:%M:%S')
        tz = build_tz(tzstring)
        dt = dt_naive.replace(tzinfo=tz)
        
        # The magical if statement, you knew it was somewhere :)
        if (now - dt) > age_limit:
            print 'Archiving message %s. Subject: %s' % (msg_id, subject)
            archive_message(s, msg_id)

def get_credentials():
    email = EMAIL_ADDRESS
    pw = PASSWORD
    if len(email) == 0:
        email = raw_input('Email address (name@gmail.com): ')
    if len(pw) == 0:
        pw = getpass.getpass()
    return email, pw

def main():
    email, password = get_credentials()
    s = connect(email, password)

    # Select inbox
    s.select('INBOX')

    # Get autoarchive:\d+ labels
    label_ages = get_autoarchive_labels(s)

    for label, age in label_ages:
        msg_ids = get_message_ids(s, label)
        archive_old_messages(s, msg_ids, age)

    # bye
    s.close()
    s.logout()
    print 'Done.'
    
    
if __name__ == '__main__':
    main()