# -*- coding:UTF-8 -*-
#
# Copyright (C) 2010 Vittorio Palmisano <vpalmisano at gmail dot com>
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
#
__author__ = u"Vittorio Palmisano  <vpalmisano at gmail dot com>"
__description__ = u"Calendar Sync 1.0\nhttp://calsync.googlecode.com"

import urllib
import httplib
import urlparse
import time
import sys
import os
sys.path.append(os.path.dirname(sys.argv[0]))
try:
    from XmlParser import XMLParser
except ImportError:
    pass

URL = 'http://www.google.com/calendar/ical/%s/private/basic.ics'

def readline(u):
    line = c = u.read(1)
    while c != '\n':
        c = u.read(1)
        if not c:
            break
        line += c
    return line

def get_google_auth(u, p):
    print 'Requesting auth...'
    params = urllib.urlencode({'Email':u, 'Passwd':p, 
        'service':'cl', 'source':'calsync-1.0'})
    headers = {"Content-type": "application/x-www-form-urlencoded", 
        "Accept": "text/plain"}
    conn = httplib.HTTPSConnection("www.google.com")
    conn.request("POST", "/accounts/ClientLogin", params, headers)
    response = conn.getresponse()
    line = readline(response)
    while line:
        try:
            k, value = line.strip().split('=', 1)
            if k == 'Auth':
                conn.close()
                return value
            if k == 'Error':
                print 'Error:', value
                conn.close()
                return ''
        except Exception:
            pass
        line = readline(response)
    conn.close()
    return ''

def get_google_calendar(auth, calid):
    print 'Opening calendar', calid
    headers = {"Accept": "text/plain", 'Authorization': 'GoogleLogin auth='+auth}
    conn = httplib.HTTPConnection("www.google.com")
    conn.request("GET", URL %calid, None, headers)
    response = conn.getresponse()
    return response

def to_time(s):
    t = [int(s[0:4]), int(s[4:6]), int(s[6:8]), 0, 0, 0, 0, 0, 0]
    if len(s) == 16:
        t[3] = int(s[9:11])
        t[4] = int(s[11:13])
        t[5] = int(s[13:15])
        return time.mktime(t)
    elif len(s) == 8:
        return time.mktime(t)
    return 0

CAL_LIST_URL = 'http://www.google.com/calendar/feeds/default/owncalendars/full'

def get_user_calendars(auth):
    print 'Opening Google calendar list...'
    headers = {"Accept": "text/xml", 'Authorization': 'GoogleLogin auth='+auth}
    conn = httplib.HTTPConnection("www.google.com")
    conn.request("GET", CAL_LIST_URL, None, headers)
    response = conn.getresponse()
    data = response.read()
    location = response.getheader('location')
    conn.request("GET", location, None, headers)
    response = conn.getresponse()
    data = response.read()
    #
    p = XMLParser()
    p.parseXML(data)
    calendars = []
    for node in p.root.childnodes['entry']:
        title = node.childnodes['title'][0].content
        link = node.childnodes['content'][0].properties['src']
        link = link.replace('http://www.google.com/calendar/feeds/', '')
        link = link.replace('/private/full', '')
        calendars.append((title, link))
    return calendars

def fetch_calendar_events(u, start=30, end=60):
    start = time.time() - 60*60*24*start
    end = time.time() + 60*60*24*end
    print 'Fetching calendar...'
    events = []
    curevent = []
    line = readline(u)
    tzoffset = 0
    while line:
        line = line.strip()
        if line.startswith('TZOFFSETFROM:+'):
            line = line.replace('TZOFFSETFROM:+', '')
            tzoffset = int(line[:2])*60*60
        if line.startswith('BEGIN:VEVENT'):
            curevent = [0, 0, '']
        if line.startswith('DTSTART'):
            if line.startswith('DTSTART;VALUE=DATE:'):
                line = line.replace('DTSTART;VALUE=DATE:', '')
            else:
                line = line.replace('DTSTART:', '')
            t = to_time(line)
            if curevent:  
                curevent[0] = t+tzoffset
        if line.startswith('DTEND'):
            if line.startswith('DTEND;VALUE=DATE:'):
                line = line.replace('DTEND;VALUE=DATE:', '')
            else:
                line = line.replace('DTEND:', '')
            t = to_time(line)
            if curevent:
                curevent[1] = t+tzoffset
        if line.startswith('SUMMARY:'):
            if curevent:
                curevent[2] = line.replace('SUMMARY:', '')
        if line.startswith('END:VEVENT'):
            if curevent:
                if curevent[0] >= start and curevent[0] <= end:
                    events.append(curevent)
                    #print ' collected:', curevent
                else:
                    break
                curevent = []
        line = readline(u)
    return events

def set_exit_if_standalone():
    appname = appuifw.app.full_name()
    print appname
    if appname[-10:] != u"Python.app":
        appuifw.app.set_exit()

try:
    import appuifw
    import calendar
    import e32
    import thread
    import socket
except ImportError:
    pass

CONFIG_FILE = u'c:\\system\\apps\\calsync.cfg'

class CalSync:
    def __init__(self):
        self.lock = e32.Ao_lock()
        appuifw.app.title = u'CalSync'
        appuifw.app.exit_key_handler = self.quit
        self.text = appuifw.Text()
        appuifw.app.body = self.text
        self.text.add(u'Application started\n')
        def list_cb():
            i = self.listbox.current()
            if i == 0:
                self.edit_account()
            elif i == 1:
                self.update_calendar()
            else:
                self.quit()
        #self.listbox = appuifw.Listbox([u'Account settings', u'Syncronize', u'Exit'], 
        #    list_cb)
        #appuifw.app.body = self.listbox
        appuifw.app.menu = [
            (u'Update from server', self.update_calendar),
            (u'Settings', (
                (u'Account', self.edit_account),
                (u'Calendars', self.get_calendars),
                (u'Access point', self.choose_ap),
                (u'Start date', self.choose_start_date),
                (u'End date', self.choose_end_date),
            )),
            (u'About', self.show_about),
            (u'Exit', self.quit),
        ]
        self.load_settings()
        if self.settings['last_update']:
            t = float(self.settings['last_update'])
            s = time.strftime('%d/%m/%Y %H:%M:%S', time.gmtime(t))
            self.text.add(u'Last update: %s\n' %(s))
        self.lock.wait()

    def show_about(self):
        appuifw.note(__description__, "info")

    def load_settings(self):
        self.settings = {'username': '', 'password': '', 'auth': '', 'ap': '',
            'calendars': '', 'last_update': '', 'start': '30', 'end': '60'}
        if not os.path.exists(CONFIG_FILE):
            return
        data = open(CONFIG_FILE, 'rb').read()
        for field in data.split('\n'):
            try:
                k,v = field.split('=', 1)
                self.settings[k] = v
            except Exception, e:
                pass
        self.set_ap()
        self.text.add(u'Settings loaded\n')

    def set_ap(self):
        if self.settings['ap']:
            for d in socket.access_points():
                if d['name'] == self.settings['ap']:
                    socket.set_default_access_point(
                        socket.access_point(d['iapid']))
                    break

    def save_settings(self):
        fp = open(CONFIG_FILE, 'wb')
        s = ''
        for k, v in self.settings.items():
            s += '%s=%s\n' %(k, str(v))
        fp.write(unicode(s, 'utf8'))
        fp.close()
        self.set_ap()
        self.text.add(u'Settings saved\n')

    def quit(self):
        self.lock.signal()
        set_exit_if_standalone()

    def choose_ap(self):
        ap = socket.select_access_point()
        for d in socket.access_points():
            if d['iapid'] == ap:
                print 'setting ap:', d['name']
                self.settings['ap'] = d['name']
                break
        self.save_settings()

    def edit_account(self):
        username = appuifw.query(u'Username (email)', 'text', 
            unicode(self.settings['username']))
        if not username:
            return
        if username != self.settings['username']:
            self.settings['auth'] = ''
        self.settings['username'] = username
        password = appuifw.query(u'Password', 'code')
        if not password:
            return
        if password != self.settings['password']:
            self.settings['auth'] = ''
        self.settings['password'] = password
        self.save_settings()

    def choose_start_date(self):
        start = appuifw.query(u'Days backward', 'number', int(self.settings['start']))
        if start != int(self.settings['start']):
            self.settings['start'] = unicode(start)    
            self.save_settings()
    
    def choose_end_date(self):
        end = appuifw.query(u'Days afterward', 'number', int(self.settings['end']))
        if end != int(self.settings['end']):
            self.settings['end'] = unicode(end)
            self.save_settings()

    def autenticate(self):
        if not self.settings['username'] or not self.settings['password']:
            appuifw.note(u"Check your account settings.", "info")
            return False
        if not self.settings['ap']:
            self.choose_ap()
        #appuifw.note(u"Logging in (%s)..." %self.settings['ap'], "info")
        self.text.add(u"Logging in (%s)...\n" %self.settings['ap'])
        if not self.settings['auth']:
            self.settings['auth'] = get_google_auth(self.settings['username'], 
                self.settings['password'])
            if not self.settings['auth']:
                appuifw.note(u"Login failed", "error")
                return
            self.save_settings()
        return True

    def get_calendars(self):
        if not self.autenticate():
            return
        self.text.add(u"Getting calendars...\n")
        calendars = get_user_calendars(self.settings['auth'])
        l = [unicode(c[0]) for c in calendars]
        index = appuifw.multi_selection_list(l , style='checkbox', search_field=0)
        c = ''
        for i in index:
            c += '%s:%s|' %(calendars[i][0], calendars[i][1])
        self.settings['calendars'] = c
        self.save_settings()

    def update_calendar(self):
        if not self.settings['calendars']:
            appuifw.note(u"Select one or more calendar to import first", "info")
            return
        if not self.autenticate():
            return
        calendars = []
        for c in self.settings['calendars'].split('|'):
            items = c.split(':', 1)
            if len(items) == 2:
                calendars.append(items)
        c = calendar.open()
        added = 0
        # get events
        for title, calid in calendars:
            self.text.add(u'Opening calendar "%s"...\n' %title)
            u = get_google_calendar(self.settings['auth'], calid)
            self.text.add(u" fetching data...\n")
            start = int(self.settings['start'])
            end = int(self.settings['end'])
            for event in fetch_calendar_events(u, start, end):
                #self.text.add(u' event: "%s"\n' %unicode(event[2], 'utf8'))
                exist = False
                for d in c.find_instances(event[0], event[1], appointments=1):
                    entry = c[d['id']]
                    if unicode(event[2], 'utf8') == entry.content:
                        exist = True
                        break
                if not exist:
                    self.text.add(u' adding: "%s"' %unicode(event[2], 'utf8'))
                    e = c.add_appointment()
                    e.set_time(event[0], event[1])
                    e.content = unicode(event[2])
                    e.commit()
                    added += 1
        appuifw.note(u"Update done (%d added)" %added, "info")
        self.text.add(u'Update done (%d added)\n' %added)
        self.settings['last_update'] = str(time.time())
        self.save_settings()

if __name__ == "__main__":
    try:
        c = CalSync()
    except Exception, e:
        print e
