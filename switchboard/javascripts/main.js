/*
 * Auto-generated content from the Brackets New Project extension.
 */

/*jslint vars: true, plusplus: true, devel: true, nomen: true, indent: 4, maxerr: 50 */
/*global $, window, document */

function timestamp() {
    var now = new Date();
    return now.toLocaleString();
}

function getLabCss(machine) {
    if (machine.lastIndexOf("B103", 0) === 0) return "b103";
    if (machine.lastIndexOf("B107", 0) === 0) return "b107";
    if (machine.lastIndexOf("B160", 0) === 0) return "b160";
    if (machine.lastIndexOf("B162", 0) === 0) return "b162";
    if (machine.lastIndexOf("B215", 0) === 0) return "b215";
    return "";
}

// Simple jQuery event handler
$(document).ready(function () {


    var pusher = new Pusher('58992988b1202996ed5e', {
        authTransport: 'jsonp',
        authEndpoint: 'http://valid-hall-624.appspot.com/pusher/auth'
    });
    // Pusher.channel_auth_endpoint = 'http://valid-hall-624.appspot.com/pusher/auth'
    var channel = pusher.subscribe('private-talk');
    
    
    
    
     channel.bind('up_event', function(data) {
    			$("#message-container").append('<li class="list-group-item" id="' + data['guid'] + '">' +
                                                  '<span class="label label-default">' + timestamp() + '</span>' +
                                                  '<span class="label ' + getLabCss(data['machine']) + '">' + data['machine'] + '</span>' +
                                                  '<span class="label label-fullname">' + data['fullName'] + '</span>' +
                                                  '</li>');
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

