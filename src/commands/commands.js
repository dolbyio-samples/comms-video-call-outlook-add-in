// 1. How to construct online meeting details.
// Set your video conference domain name as an environment variable either in your local .env file or in your deployment environment
// Or you can add it  as a hardcoded string for testing purposes.
const domain = process.env.DOMAIN_NAME;

// Generate an instant meeting UID
const meetingID = crypto.randomUUID();

const meetingLink = domain+ meetingID;
const QRGenerator = "https://chart.googleapis.com/chart?chs=150x150&cht=qr&choe=UTF-8&chl=";
const newBody = '<br>' +
    'To join the Dolby.io video call on a computer or mobile phone:'+
    '<br><br>'+
    `<a href=${meetingLink} target="_blank">${meetingLink}</a>` +
    '<br><br>' +
    'or scan the QR code below:'+
    '<br>'+
    `<img src= ${QRGenerator+meetingLink}>`+
    '<br><br>'+
    `Meeting ID: ${meetingID}` +
    '<br><br>' +
    'Want to test your video connection?' +
    '<br>'+
    `<a href=${domain} target="_blank">Join a test meeting</a>`;
    

// 2. How to define and register a function command named `insertDolbyioMeeting` (referenced in the manifest)
//    to update the meeting body with the online meeting details.
function insertDolbyioMeeting(event) {
    // Get HTML body from the client. Uses updateBody to append at the end of existing body.
    // mailboxItem.body.getAsync("html",
    //     { asyncContext: event },
    //     function (getBodyResult) {
    //         if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
    //             updateBody(getBodyResult.asyncContext, getBodyResult.value);
    //         } else {
    //             console.error("Failed to get HTML body.");
    //             getBodyResult.asyncContext.completed({ allowEvent: false });
    //         }
    //     }
    // );
    // Sets body from scratch each time the add-in works.
        mailboxItem.body.setAsync(newBody,
            { asyncContext: event, coercionType: "html" },
            function (setBodyResult) {
                if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    setBodyResult.asyncContext.completed({ allowEvent: true });
                } else {
                    console.error("Failed to set HTML body.");
                    setBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    
        setLocation(event, meetingLink);
}

// 3. How to implement a supporting function `updateBody`
//    that appends the online meeting details to the current body of the meeting.
function updateBody(event, existingBody) {
    // Append new body to the existing body.
    mailboxItem.body.setAsync(existingBody + newBody,
        { asyncContext: event, coercionType: "html" },
        function (setBodyResult) {
            if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                setBodyResult.asyncContext.completed({ allowEvent: true });
            } else {
                console.error("Failed to set HTML body.");
                setBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );
}
function setLocation(event,meetingLink){
    mailboxItem.location.setAsync(meetingLink,  
        { asyncContext: event },(result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${result.error.message}`);
        result.asyncContext.completed({ allowEvent: false });
        return;
        }
        console.log(`Successfully set location to ${meetingLink}`);
        result.asyncContext.completed({ allowEvent: true });
    });
}
let mailboxItem, userName;
// Office is ready.
Office.onReady(function () {
    mailboxItem = Office.context.mailbox.item;
}
);
// Register the function.
Office.actions.associate("insertDolbyioMeeting", insertDolbyioMeeting);
