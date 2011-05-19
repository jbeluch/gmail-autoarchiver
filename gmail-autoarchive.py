#!/usr/bin/env python
# Copyright (C) 2011 by Jonathan Beluch
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
'''
GMail Auto-Archiver

This script will archive emails in your inbox that are older than a
specified number of days.

Usage:
    1. Set up filters in gmail matching the pattern 'aa:\d+' where the
       '\d+' is the age limit in days, e.g. 'aa:3'.
    2. [Optional] Enter you email address below for the variable
       EMAIL_ADDRESS. If you skip this step, you will be prompted to
       enter your email address interactively when the script runs.
    3. Execute the script. If you are running it for the first time,
       you will be required to authorize the script's oauth access in
       a web browser.

    The script will print the subject line for any emails that are
    auto-archived.

    Note: Once you authorize the oauth token/secret, they are saved to
    disk at OAUTH_PATH. If the token/secret no longer work, simply
    remove the file at OAUTH_PATH. The next time the script is run it
    will set up a new token/secret.

Further Reading:
    GMail IMAP: http://code.google.com/apis/gmail/imap/
    IMAP Protocol: http://tools.ietf.org/search/rfc3501
    python imaplib: http://docs.python.org/library/imaplib.html

Hints about GMail:
    - GMail labels are = to IMAP folders.
    - To archive an email in gmail when you have seleted INBOX, you
      simply set the IMAP +FLAG \Deleted. This will remove it from the 
      INBOX folder but the ALL MAIL folder will still contain a copy.
'''

'''next features
- oauth instead of pw in file *completed*
- download all labels and use regex to match instead of autoarchive:*
- fetch all messages at once instead of using a separate request for each.
'''

from datetime import datetime, tzinfo, timedelta
import imaplib
import email
from lib import xoauth
from itertools import chain

## Config -------------------------------------------------------------

# Either hard code values here, or leave blank and you will be prompted
# upon script execution.
EMAIL_ADDRESS = '' # yourname@gmail.com

# Where the oauth token/secret are stored.
# Remove this file and run script to generate a new token/secret.
OAUTH_PATH = '.oauth_identity'

# Gmail label pattern, * is a wildcard
# This pattern must conform to the IMAP spec listed here:
#   http://tools.ietf.org/search/rfc3501#section-6.3.8
# Typically it will be something like 'autoarchive:*'
# Note that later in the script the labels are expected to contain
# a color and then an integer reprsenting the number of days after the
# colon.
LABEL_PATTERN = 'aa:*'

## End Config ---------------------------------------------------------

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

def connect(oauth_entity, email):
    consumer = xoauth.OAuthEntity('anonymous', 'anonymous')
    xoauth_string = xoauth.GenerateXOauthString(
        consumer, oauth_entity, email, 'imap',
        None, None, None)

    imap_conn = imaplib.IMAP4_SSL('imap.gmail.com')
    #imap_conn.debug = 4
    imap_conn.authenticate('XOAUTH', lambda x: xoauth_string)
    
    print 'Connected to mailbox successfully.'
    return imap_conn

def get_autoarchive_labels(s, label_pattern):
    '''Returns a list of tuples (str labelname, int age_in_days)'''
    _, list_of_labels = s.list(pattern=label_pattern)
    # Annoyingly if has no matches it returns a list with one element, None
    if list_of_labels[0] == None:
        return []

    # labels looks like: '(\\HasNoChildren) "/" "aa:1"'
    # we want to extract the 'aa:1' part
    ret = []
    #print list_of_labels
    for item in list_of_labels:
        label = item.split('"')[-2] # label = 'aa:3'
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

def fetch_emails(s, msg_ids):
    '''Returns a dict, keys are msg_ids, values are email objects.
    Currently hardcoded to only fetch the date and subject headers.'''
    msg_id_str = ','.join(msg_ids)
    _, messages = s.fetch(msg_id_str, '(body[header.fields (date subject)])')

    # Every 2nd item is a closing ')' so we skip by 2
    emails = {}
    for msg in messages[::2]:
        names, values = msg
        msg_id, _ = names.split(' ', 1)
        emails[msg_id] = email.message_from_string(values)
    return emails

def get_messages_to_archive(ages, emails):
    '''Returns a list of msg ids to be archived.'''
    # ages = {msg_id: age, ... }
    # emails = {msg_id: email, ... }
    now = datetime.utcnow().replace(tzinfo=utc)

    old_msgs = []
    for msg_id, mail in emails.items():
        datestr = mail.get('Date').strip()
        # e.g. datestr = 'Thu, 7 Apr 2011 08:34:04 -0400 (EDT)'
        subject = mail.get('Subject').replace('\r\n', ' ')

        # %z doesn't work , must manually build timezone.
        # We want to remove the offset and sometimes present tzname from 
        # datestr so we can use strptime() 
        
        # sometimes datestr has '(tzname)' at the end and sometimes not, so we
        # split from left a set amount, guess i could use re
        parts = datestr.split(' ', 5)
        datetimestr = ' '.join(parts[:5])
        tzstring = parts[5]

        # Make dt object and apply the correct tzinfo
        dt_naive = datetime.strptime(datetimestr, r'%a, %d %b %Y %H:%M:%S')
        tz = build_tz(tzstring)
        dt = dt_naive.replace(tzinfo=tz)
        
        assert msg_id in ages.keys(), 'No age limit for message %s' % msg_id
        age_limit = timedelta(days=ages[msg_id])

        # The magical if statement, you knew it was somewhere :)
        if (now - dt) > age_limit:
            print 'Preparing message %s to be archived. Subject: %s' % (msg_id, subject)
            old_msgs.append(msg_id)

    return old_msgs

def archive_messages(s, msg_ids):
    ''' Simply set the deleted flag and msg will be archived in 
    gmail. '''
    print 'Archiving messages.'
    msg_str = ','.join(msg_idss)
    s.store(msg_str, '+FLAGS', '"\\\\Deleted"')


def ask_for_email():
    email = raw_input('Email address (name@gmail.com): ')
    return email.strip()

def write_oauth_identity(fn, oauth_entity):
    with open(fn, 'w') as f:
        f.write('%s\n%s' % (oauth_entity.key, oauth_entity.secret))

def read_oauth_identity(fn):
    '''Returns an xoauth.OAuthEntity from a given fn.
    Expects the token on line 1 and the secret on line 2.'''
    try:
        with open(fn) as f:
            lines = f.readlines()
    except IOError:
        print 'Error reading from %s. Check that it exists.' % fn
        return None
    return xoauth.OAuthEntity(lines[0].strip(), lines[1].strip())

def generate_new_oauth_entity(user, fn):
    '''Generates a new oauth token/secret for a given user. The new
    pair is written to the given fn. '''
    scope = 'https://mail.google.com/'
    consumer = xoauth.OAuthEntity('anonymous', 'anonymous')
    google_accounts_url_generator = xoauth.GoogleAccountsUrlGenerator(user)
    request_token = xoauth.GenerateRequestToken(consumer, scope, None, None, google_accounts_url_generator)

    # Wait for user to visit URL and authenticate this application.  After
    # authenticating this application, they must paste the verification code.
    oauth_verifier = raw_input('Enter verification code: ').strip()

    # Get the token and token secret to be saved and used going forward for
    # authentiaction
    access_token = xoauth.GetAccessToken(consumer, request_token, oauth_verifier,
                          google_accounts_url_generator)

    if not access_token:
        print 'There was a problem getting a valid access token.'
        return None

    # Save the credentials for the future.
    write_oauth_identity(fn, access_token)

    return access_token

def main():
    # Check if email is stored otherwise we have to ask for it
    if len(EMAIL_ADDRESS) == 0:
        email = ask_for_email()

    # First check for oauth credentials
    oauth_entity = read_oauth_identity(OAUTH_PATH)

    # If not saved credentials, attempt to create new ones
    if not oauth_entity:
        print 'No existing OAuth credentials found. Attempting to create a new identity.'
        oauth_entity = generate_new_oauth_entity(email, OAUTH_PATH)

    # If the user mistypes the verification code, generate_new_oauth_entity
    # will return None
    if not oauth_entity:
        print 'Cannot continue without a valid token.'
        return

    # Connect to the server using oauth
    s = connect(oauth_entity, email)

    # Select inbox
    s.select('INBOX')

    # Get aa:\d+ labels
    label_ages = get_autoarchive_labels(s, LABEL_PATTERN)

    # Create a dict, keys are msg_ids, values are the age_limit parsed from the
    # gmail label
    ages = {}
    for label, age in label_ages:
        msg_ids = get_message_ids(s, label)
        ages.update([(msg_id, age) for msg_id in msg_ids])


    # Get the message headers for all messages in question with a single FETCH
    all_msg_ids = ages.keys()
    emails = fetch_emails(s, all_msg_ids)

    # Get a message ids for emails to be archived based on email date
    old_msgs = get_messages_to_archive(ages, emails)

    if len(old_msgs) > 0:
        archive_messages(s, old_msgs)
    else:
        print 'No messages to be archived.'

    # bye
    s.close()
    s.logout()
    
    
if __name__ == '__main__':
    main()
    print 'Done.'
