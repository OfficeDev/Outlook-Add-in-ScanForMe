
/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            detectActionsForMe();
        });
    };

    function detectActionsForMe() {
        var myInfo = Office.context.mailbox.userProfile;

        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {

            // Sent to me
            for (var i = 0; i < item.to.length; i++)
            {
                if (item.to[i].emailAddress == myInfo.emailAddress)
                {
                    $("#to-result").text("yes");
                    break;
                }
            }

            // I was cc'd
            for (var i = 0; i < item.cc.length; i++) {
                if (item.cc[i].emailAddress == myInfo.emailAddress) {
                    $("#cc-result").text("yes");
                    break;
                }
            }

            Office.context.mailbox.item.body.getAsync(function (asyncResult) {
                var bodyText = asyncResult.value;
                var nameParts = myInfo.displayName.split(" ");
                
                // Make an assumtion that the displayName is in the form "firstName lastName"
                var myFirstName = nameParts[0];

                // Create regular expression to find all matches of the first name
                // i => ignore case
                // g => global match, i.e., doesn't stop after first match
                var re = new RegExp(myFirstName, 'gi'); 
                var results = new Array();
                while (re.exec(bodyText)) {
                    results.push(re.lastIndex);
                }

                
                if (results.length > 0)
                {
                    $("#body-result").text("yes");
                }
                
            });
            // Indirectly sent to me
            // My name is in the body (first name)
            // How urgent is the email?
            // Anything highlighted?
           
        }
    }

})();

// *********************************************************
//
// Outlook-Add-in-ActionDetector, https://github.com/OfficeDev/Outlook-Add-in-ActionDetector
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