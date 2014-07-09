/*
 * Auto-generated content from the Brackets New Project extension.
 */

/*jslint vars: true, plusplus: true, devel: true, nomen: true, indent: 4, maxerr: 50 */
/*global $, window, document */

// Simple jQuery event handler
$(document).ready(function () {


    var pusher = new Pusher('58992988b1202996ed5e', {
        authTransport: 'jsonp',
        authEndpoint: 'http://valid-hall-624.appspot.com/pusher/auth'
    });
    // Pusher.channel_auth_endpoint = 'http://valid-hall-624.appspot.com/pusher/auth'
    var channel = pusher.subscribe('private-talk');
    
    
     channel.bind('up_event', function(data) {
    			$("#message-container ol").append('<li id="' + data['guid'] + '">' + data['msg'] + '</li>');
    		});
    
    channel.bind('down_event', function(data) {
    			$('#' + data['guid']).remove();
    		});
    
    channel.bind('pusher:subscription_error', function(status) {
           alert("Unable to Listen!");
    });
    
    channel.bind('pusher:subscription_succeeded', function() {
        alert("Successfully Listening!");
    });
    
});

