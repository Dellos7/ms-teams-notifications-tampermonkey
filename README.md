# Microsoft Teams Browser Notifications
Tampermonkey script in order to enable browser notifications for the Web-based Microsoft Teams chat. Useful in Linux (in Linux chat notifications do not work).

Tested and working:
* Chrome 66 / Ubuntu 16.04
* Chrome 68 / Windows 10

## Installation

1. Install Chrome's [Tampermonkey extension](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo?hl=es).
2. Install [Greasy Fork installation script](https://greasyfork.org/es/scripts/48399-microsoft-teams-notifications)
3. Reload https://teams.microsoft.com/
4. **Enable** browser notifications

![https://image.ibb.co/ig5GJ9/teams_notifications.png](https://image.ibb.co/ig5GJ9/teams_notifications.png)

## How does it work

What the script basically does is to listen to the `<title>` tag on the `<head>` tag of the html page. If that title changes (e.g before: `Chat | Mom`, after: `(1) Chat | Dad`) then the script sends a browser notification in order to alert the user that there's a new notification in the teams application.

There is also other feature that some seconds after the user has left the teams tab in the browser, it focuses on the **T-Bot** chat. This is actually necessary because **the notifications do not work if you receive some message from the chat you are currently focused on**.

## Roadmap

1. Add lang support (currently messages are written on spanish).
