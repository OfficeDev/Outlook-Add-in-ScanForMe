/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};

    // Common initialization function (to be called from each page)
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding ms-bgColor-themeDarkAlt">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header" class="ms-font-m ms-fontColor-white"></div>' +
                    '<div id="notification-message-body" class="ms-font-m ms-fontColor-white"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });


        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };
    };

    return app;
})();