#!/usr/bin/env python
#

import logging
import datetime
import pusher

from google.appengine.ext import webapp

try:
  import config
except ImportError:
  raise Exception("refer to the README on creating a config.py file")

    
class MainHandler(webapp.RequestHandler):
    def post(self):
        
        machine_name = self.request.get('machine')
        user_name = self.request.get('userName')
        full_name = self.request.get('fullName')
        dest_state = self.request.get('state')
        guid = self.request.get('guid')
    
        p = pusher.Pusher(app_id=config.app_id, key=config.app_key, secret=config.app_secret)
        
        event_data = {'guid': guid, 'msg': machine_name + ':' + user_name, 'userName' : user_name, 'state': dest_state, 'machine': machine_name}
        
        if dest_state == "Up":
            p['private-talk'].trigger('up_event', event_data)
        else:
            p['private-talk'].trigger('down_event',event_data)
        now = datetime.datetime.now().strftime("%c")
        logging.info(guid + ":" + dest_state + ":" + now + ":" + machine_name + ":" + user_name + ":" + full_name)

app = webapp.WSGIApplication([('/', MainHandler)],
                                         debug=True)
