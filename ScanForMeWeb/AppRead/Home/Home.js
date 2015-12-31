/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path='../App.js' />

(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            detectActionsForMe();
        });
    };

    function detectActionsForMe() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);

        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            var myInfo = Office.context.mailbox.userProfile;

            // If the message is from me, no point in checking any further
            if (item.from.emailAddress === myInfo.emailAddress) {
                app.showNotification('Mail Not Scanned', 'It\'s from you, so we didn\'t think there was any point.');
            }
            else {
                var nameParts = myInfo.displayName.split(' ');

                // Make an assumption that the displayName is in the form 'firstName lastName'
                var myFirstName = nameParts[0];

                // Check whether I was cc'd on this email
                for (var i = 0; i < item.cc.length; i++) {
                    if (item.cc[i].emailAddress === myInfo.emailAddress) {
                        $('#cc-result').text('yes');
                        break;
                    }
                }

                // Check whether I am on the To line
                for (var i = 0; i < item.to.length; i++) {
                    if (item.to[i].emailAddress === myInfo.emailAddress) {
                        $('#to-result').text('yes');
                        break;
                    }
                }

                // We need to determine if body.getAsync() is defined. We require this method in order to
                // retrieve the body test for parsing. This method was added in v1.3 of the API
                // and may not be available on every Outlook client.
                //
                // For more information, please see Understanding API Requirement Sets at
                // https://dev.outlook.com/reference/add-ins/tutorial-api-requirement-sets.html
                if (Office.context.mailbox.item.body.getAsync !== undefined) {
                    // Check whether I am mentioned in the body of the email by name
                    // In this sample we scan the email body as plain text. You can also
                    // set the coercionType on the getAsync() method to retrieve the body as HTML.
                    // For an example of retrieving the body as HTML and parsing the result,
                    // see https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js
                    Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                        var bodyText = asyncResult.value;

                        // Create regular expression to find all matches of the first name
                        // i => ignore case
                        // g => global match, i.e., doesn't stop after first match
                        var regex = new RegExp(myFirstName, 'gi');
                        var matchingArray = new Array();
                        while (regex.exec(bodyText)) {
                            matchingArray.push(regex.lastIndex);
                        }

                        var result = 'Scan Complete.';

                        if (matchingArray.length > 0) {
                            app.showNotification('Scan Complete', 'It looks like you are mentioned by name in the body of this email');
                        }
                        else {
                            app.showNotification('Scan Complete', 'It looks like you are not mentioned by name in the body of this email');
                        }

                    });
                }
                else { // Method not available
                    app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
                }
            }
        }
    }
})();

// *********************************************************
//
// Outlook-Add-in-ScanForMe, https://github.com/OfficeDev/Outlook-Add-in-ScanForMe
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************