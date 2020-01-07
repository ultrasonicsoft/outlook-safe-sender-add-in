/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

const VALID_DOMAIN_NAME = 'gmail.com';
const SUCCESS_MESSAGE = 'Sender is safe!';
const ERROR_MESSAGE = 'WARNING!!! SENDER IS OUTSIDER!';

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
    const sender = Office.context.mailbox.item.sender.emailAddress.split('@')[1];
    console.log(sender);
    const isSafe = sender === VALID_DOMAIN_NAME;

    const errorMessage: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: ERROR_MESSAGE
    };

    const successMessage: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: SUCCESS_MESSAGE,
        icon: "Icon.80x80",
        persistent: true
    };

    // Show a notification message
    Office.context.mailbox.item.notificationMessages.replaceAsync("action", isSafe ? successMessage : errorMessage);

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
            ? window
            : typeof global !== "undefined"
                ? global
                : undefined;
}

const g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.action = action;
