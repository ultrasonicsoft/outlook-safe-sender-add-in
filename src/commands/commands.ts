/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
    // If needed, Office.js is ready to be called

    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, selectedItemChanged);
    console.log('itemchanged event subscribed...');
});

function selectedItemChanged(eventArgs: any) {
    console.log('itemchanged event fired....');
    console.log(eventArgs);
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
    const sender = Office.context.mailbox.item.sender.emailAddress.split('@')[1];
    console.log(sender);
    const isSafe = sender === 'boehringer-ingelheim.com';

    const errorMessage: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: 'Sender is outsider!!!'
    };

    const successMessage: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Sender is safe. ",
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
