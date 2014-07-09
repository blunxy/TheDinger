#!/usr/bin/env python
#
from google.appengine.ext import webapp
import pusher
try:
    import json
except ImportError:
    import simplejson as json
    
try:
  import config
except ImportError:
  raise Exception("refer to the README on creating a config.py file")

class AuthHandler(webapp.RequestHandler):
    def get(self):
    
      channel_name = self.request.get('channel_name')
      socket_id = self.request.get('socket_id')
      callback = self.request.get('callback')
    
      p = pusher.Pusher(app_id=config.app_id, key=config.app_key, secret=config.app_secret)
    
      auth = p[channel_name].authenticate(socket_id)
    
      
      json_data = json.dumps(auth)
    
      self.response.headers['Content-Type'] = "application/json"
      self.response.out.write(callback + '(' + json_data +')')


app = webapp.WSGIApplication([('/pusher/auth', AuthHandler)],
                                         debug=True)
