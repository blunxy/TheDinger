/*
 * Auto-generated content from the Brackets New Project extension.
 */

/*jslint vars: true, plusplus: true, devel: true, nomen: true, indent: 4, maxerr: 50 */
/*global $, window, document */

// Simple jQuery event handler
$(document).ready(function () {
    "use strict";
    $("#clickMe").click(function (e) {
        alert("hello world!");
    });


    var pusher = new Pusher('58992988b1202996ed5e', {
        authTransport: 'jsonp',
        authEndpoint: 'http://valid-hall-624.appspot.com/pusher/auth'
    });
    // Pusher.channel_auth_endpoint = 'http://valid-hall-624.appspot.com/pusher/auth'
    var channel = pusher.subscribe('private-talk');
    
    
     channel.bind('my_event', function(data) {
    			alert(data.msg);
    		});
    
    channel.bind('pusher:subscription_error', function(status) {
        if (status === 404) {
            console.log("404");
        } else if (status === 408) {
            
        } 
    });
    channel.bind('pusher:subscription_succeeded', function() {
        alert("Bingo!");
    });
    
});

