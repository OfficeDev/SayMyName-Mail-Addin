(function(){
  'use strict';

  angular.module('officeAddin')
         .controller('homeController', ['$scope', '$http', 'dataService', homeController]);

  /**
   * Controller constructor
   */
  function homeController($scope, $http, dataService){
    
    //utility function for checking for email in array
    var emailExists = function (array, email) {
      for (var i = 0; i < array.length; i++) {
        if (array[i].email == email)
          return true;
      }
      return false;
    };
            
    //utility for getting all appointment participants
    var getAppointmentParticipants = function (appointment) {
      //TODO: check type to ensure it is a user and not DL???????
      var participants = [];
      
      if (appointment.requiredAttendees) {
        for (var i = 0; i < appointment.requiredAttendees.length; i++) {
          if (!emailExists(participants, appointment.requiredAttendees[i].emailAddress.toLowerCase()))
            participants.push({
              email: appointment.requiredAttendees[i].emailAddress.toLowerCase(),
              name: appointment.requiredAttendees[i].displayName
            });
        }
      }
        
      if (appointment.optionalAttendees) {
        for (var i = 0; i < appointment.optionalAttendees.length; i++) {
          if (!emailExists(participants, appointment.optionalAttendees[i].emailAddress.toLowerCase()))
            participants.push({
              email: appointment.optionalAttendees[i].emailAddress.toLowerCase(),
              name: appointment.optionalAttendees[i].displayName
            });
        }
      }
          
      return participants;
    };

    var getParticipants = function () {
      //TODO: check type to ensure it is a user and not DL???????
      var participants = [];
  
      //get the mailbox user off the mailbox userProfile
      var email = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
      participants.push({
        email: email,
        name: Office.context.mailbox.userProfile.displayName
      });
  
      //process the SENDER
      var frm = Office.cast.item.toMessageRead(outlookItem).from;
      if (!emailExists(participants, frm.emailAddress.toLowerCase()))
        participants.push({
          email: frm.emailAddress.toLowerCase(),
          name: frm.displayName
        });
  
      //process TO
      var to = Office.cast.item.toMessageRead(outlookItem).to;
      for (var i = 0; i < to.length; i++) {
        if (!emailExists(participants, to[i].emailAddress.toLowerCase()))
          participants.push({
            email: to[i].emailAddress.toLowerCase(),
            name: to[i].displayName
          });
      }
  
      //process CC
      var cc = Office.cast.item.toMessageRead(outlookItem).cc;
      for (var i = 0; i < cc.length; i++) {
        if (!emailExists(participants, cc[i].emailAddress.toLowerCase()))
          participants.push({
            email: cc[i].emailAddress.toLowerCase(),
            name: cc[i].displayName
          });
      }
  
      return participants;
    };
    
    //capture information off the item
    var outlookItem = Office.cast.item.toItemRead(Office.context.mailbox.item);
    var mailbox = Office.context.mailbox;
    
    if (outlookItem.itemType === "message") {
      $scope.people = getParticipants();
    }
    else {
      $scope.people = getAppointmentParticipants();
    }
    
    $scope.waiting = false;
    $scope.activeAudio = null;
    $scope.listen = function(index) {
      $scope.waiting = true;
      if ($scope.people[index].audio) {
        $scope.activeAudio = $scope.people[index].audio;
        $("#mp3audio").attr("src", $scope.people[index].audio);
        var audio = document.getElementById("mp3audio");
        audio.load();
        audio.play();
        $scope.waiting = false;
      }
      else {
        $http.post("https://speechapis.azurewebsites.net/api/speech/", JSON.stringify($scope.people[index].name))
          .success(function (data) {
            $scope.people[index].audio = data;
            $scope.activeAudio = data;
            
            $("#mp3audio").attr("src", data);
            var audio = document.getElementById("mp3audio");
            audio.load();
            audio.play();
            $scope.waiting = false;
          })
          .error(function (err) {
            //TODO
            var x = "";
          });
      }
    }
  }

})();
