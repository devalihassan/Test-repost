var app = angular.module('SAPApp', ['ngMaterial'], function ($mdThemingProvider) {

    $mdThemingProvider.theme('default')
        .primaryPalette('blue', {
            'default': '400', // by default use shade 400 from the pink palette for primary intentions

        })

//pro
});
app.controller('AppCtrl', function ($scope, $mdDialog, $mdToast, $log, $window, $http) {

    $scope.isUserLoggedIn = true;
   
    var settingDialoge;
    var loginDialoge;
    var uploadDialoge;

    let emailbase64;


    Office.onReady(function (info) {



        if (info.host === Office.HostType.Outlook) {
            // Check if the host application is Outlook
            var checkLogin = localStorage.getItem('login');
            if (checkLogin === "true") {
                $scope.isUserLoggedIn = false;
                if (!$scope.$$phase) {
                    $scope.$apply();
                };
            }
            else {
                openSetting();
            };
           
        };





        function openSetting() {
            Office.context.ui.displayDialogAsync('https://localhost:44349/Templates/settings.html', { height: 25, width: 40 },
                function (asyncResult) {
                    settingDialoge = asyncResult.value;
                    settingDialoge.addEventHandler(Office.EventType.DialogMessageReceived, processMessageSetting);
                }
            );
        }

        function processMessageSetting(arg) {
            console.log(arg)
            settingDialoge.close();
            if (arg.message === "submit") {
                openfiledilog();
            };
        };




        function openfiledilog() {
            Office.context.ui.displayDialogAsync('https://localhost:44349/Templates/login.html', { height: 50, width: 30 },

                function (asyncResult) {
                    loginDialoge = asyncResult.value;
                    loginDialoge.addEventHandler(Office.EventType.DialogMessageReceived, processMessageLogin);
                }
            );
        };

        function processMessageLogin(arg) {
          
            console.log(arg)


            if (arg.message == "True") {
                $scope.isUserLoggedIn = false;
                loginDialoge.close();

                if (!$scope.$$phase) {
                    $scope.$apply();
                };
               
            }
        };

     

        var filename = Office.context.mailbox.item.subject;

        $scope.saveSelMail = function () {

            window.localStorage.setItem("mailSubject", Office.context.mailbox.item.subject);

            console.log(filename);

            Office.context.ui.displayDialogAsync('https://localhost:44349/Templates/uploadFile.html', { height: 30, width: 40 },

                function (asyncResult) {
                    uploadDialoge = asyncResult.value;
                    uploadDialoge.addEventHandler(Office.EventType.DialogMessageReceived, processMessageUpload);
                }
            );


        };

        function processMessageUpload(arg) {
            // Close the upload dialog (if it's supposed to be closed here)
            uploadDialoge.close();

            if (arg.message === "cancel") {
                // Handle the "cancel" case
                console.log("Message upload canceled");
            } else {
                // Close the dialog here if it's not supposed to be closed above

                Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
                    if (result.status === "succeeded") {
                        var token = result.value;

                        // Define the filename (replace with the actual filename)
                        var filename = "example.txt";

                        var getMessageUrl = Office.context.mailbox.restUrl +
                            '/v2.0/me/messages/' + getItemRestId() + '/$value';

                        fetch(getMessageUrl, {
                            method: 'GET',
                            headers: {
                                Authorization: 'Bearer ' + token
                            }
                        }).then(function (response) {
                            if (response.ok) {
                                response.blob().then(function (blob) {
                                    var reader = new FileReader();
                                    reader.readAsDataURL(blob);
                                    reader.onloadend = function () {
                                        var base64data = reader.result;
                                        var base64 = base64data.split(',')[1];

                                        if (base64) {
                                            var form = new FormData();
                                            form.append("originalName", filename);
                                            form.append("matterId", arg.message);
                                            form.append("file", base64);
                                            form.append("USER_NAME", "TestUser");
                                            // Make an AJAX request to upload the attachment
                                            $.ajax({
                                                url: "https://grazingdelights.com.au/LPDM/RT/WS/uploadMatterAttachment",
                                                type: "POST",
                                                data: form,
                                                processData: false,
                                                contentType: false,
                                                success: function (response) {
                                                    console.log(response);
                                                },
                                                error: function (error) {
                                                    console.error(error);
                                                }
                                            });
                                        }
                                    };
                                });
                            } else {
                                console.error("Failed to fetch the email message");
                            }
                        });
                    } else {
                        console.error("Failed to get a callback token");
                    }
                });
            }
        }

        // Function to get the item REST ID (replace with your logic)
     
        function getItemRestId() {
            if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
                // itemId is already REST-formatted.
                return Office.context.mailbox.item.itemId;
            } else {
                // Convert to an item ID for API v2.0.
                return Office.context.mailbox.convertToRestId(
                    Office.context.mailbox.item.itemId,
                    Office.MailboxEnums.RestVersion.v2_0
                );
            }
        }

        $scope.GetWholeMail = function () {



           

        };
        $scope.logout = function () {

            window.localStorage.removeItem('login');
            $scope.isUserLoggedIn = true;
            openSetting()
        };



    });


});