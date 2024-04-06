
(function () {
    "use strict";

    var messageBanner;

    Office.initialize = function (reason) {

        $(document).ready(function () {

            var element = document.querySelector('.MessageBanner');

            messageBanner = new components.MessageBanner(element);

            messageBanner.hideBanner();



            // Registers an event handler to identify when messages are selected.
            Office.context.mailbox.addHandlerAsync(Office.EventType.RecipientsChanged, writeGreetingsForRecipients, (asyncResult) => {

                if (asyncResult.status === Office.AsyncResultStatus.Failed) {

                    showNotification('Error!', asyncResult.error.message);

                    return;
                }

                console.log("Event handler added for the RecipientsChanged event.");
            });

            writeGreetingsForRecipients();
        });
    };

    // Format an EmailAddressDetails object as
    // GivenName Surname <emailaddress>
    function buildEmailAddressString(address) {
        return address.displayName + " &lt;" + address.emailAddress + "&gt;";
    }

    // Take an array of EmailAddressDetails objects and
    // build a list of formatted strings, separated by a line-break
    function buildEmailAddressesString(addresses) {
        if (addresses && addresses.length > 0) {
            var returnString = "";

            for (var i = 0; i < addresses.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + buildEmailAddressString(addresses[i]);
            }

            return returnString;
        }

        return "None";
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }


    // writes greeting message for the recipietns into email body
    function writeGreetingsForRecipients() {

        Office.context.mailbox.item.to.getAsync(function (asyncResult) {
            //recepients
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                const msgTo = asyncResult.value;

                const userName = buildEmailAddressString(msgTo[i].emailAddress);

                const greetingMsg = `Hi ${userName},`;

                Office
                    .context
                    .mailbox
                    .item
                    .body
                    .setAsync(
                        greetingMsg,
                        {
                            coercionType: "html",
                            asyncContext: "123"
                        },
                        function callback(result) {

                            console.log('From inserting into body callback');
                        });
            }
            else {

                showNotification('Error!', asyncResult.error);
            }
        })
    };
})();