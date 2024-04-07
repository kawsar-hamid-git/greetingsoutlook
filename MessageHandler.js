
(function () {
    "use strict";

    var messageBanner;

    Office.initialize = function (reason) {

        $(document).ready(function () {

            var element = document.querySelector('.MessageBanner');

            messageBanner = new components.MessageBanner(element);

            messageBanner.hideBanner();



            // Registers an event handler to identify when messages are selected.
            Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecipientsChanged, writeGreetingsForRecipients, (asyncResult) => {

                if (asyncResult.status === Office.AsyncResultStatus.Failed) {

                    showNotification('Error!', asyncResult.error.message);

                    return;
                }

                console.log("Event handler added for the RecipientsChanged event.");
            });

            writeGreetingsForRecipients();
        });
    };

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

                if (msgTo.length == 0) {

                    return;
                }

                const userName = msgTo[0].displayName || msgTo[0].emailAddress;

                const nameMatching = userName.match(/^([^@]*)@/);

                userName = nameMatching ? nameMatching[1] : userName;

                const greetingMsg = `Hi ${userName},`;

                Office
                    .context
                    .mailbox
                    .item
                    .body
                    .setSelectedDataAsync(
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