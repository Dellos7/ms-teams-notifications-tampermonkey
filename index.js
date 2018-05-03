// ==UserScript==
// @name         Microsoft Teams Notifications
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  Creates browser notifications for the Web-based Teams application. Useful in Linux (in Linux notifications do not work). Tested in Chrome 66.
// @author       David López Castellote
// @match        https://teams.microsoft.com/*
// @grant        none
// ==/UserScript==
(function() {
    'use strict';

    function notifyMe(numberOfMessages) {
        // Let's check if the browser supports notifications
        if (!("Notification" in window)) {
            alert("Este navegador no soporta notificaciones de escritorio.");
        }

        // Let's check whether notification permissions have already been granted
        else if (Notification.permission === "granted") {
            // If it's okay let's create a notification
            createNotification(numberOfMessages);
        }

        // Otherwise, we need to ask the user for permission
        else if (Notification.permission !== "denied") {
            Notification.requestPermission(function(permission) {
                // If the user accepts, let's create a notification
                if (permission === "granted") {
                    createNotification(numberOfMessages);
                }
            });
        }

    }

    function createNotification(numberOfMessages) {
        var title = "Teams Online";
        var options = {
            body: "Tienes " + numberOfMessages + " nuevas notificaciones.",
            icon: document.querySelector('link[rel="icon"]').href,
            requireInteraction: true
        };
        var notification = new Notification(title, options);
        notification.onclick = function() {
            window.focus();
        };
    }


    function setTitleObserver() {
        console.log('Activando notificaciones de Teams...');
        requestNotificationsPermission();
        var target = document.querySelector('head > title');
        var observer = new window.WebKitMutationObserver(function(mutations) {
            mutations.forEach(function(mutation) {
                var newTitle = mutation.target.textContent;
                var res;
                try {
                    res = newTitle.match(/(?<=\().+?(?=\))/);
                }
                catch( e ) {
                    res = newTitle.match(/\(([^)]+)\)/);
                }
                if (res && res[0] && res[0] && document.hidden) {
                    notifyMe(res[0]);
                    return false;
                }
            });
        });
        observer.observe(target, {
            subtree: true,
            characterData: true,
            childList: true
        });

    }

    function requestNotificationsPermission() {
        Notification.requestPermission().then(function(result) {
            console.log('Permiso para notificaciones de Teams: ' + result);
        });
    }

    function exitFromCurrentChat() {
        setInterval( function() {
            if( document.hidden ) {
                document.querySelector('ul[data-tid="active-chat-list"] li:nth-last-child(2) a').click();
            }
        }, 30000 );
    }

    setTitleObserver();
    exitFromCurrentChat();
})();
