/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
var gItem;
$(document).ready(function () {
  document.getElementById("logger").innerHTML += "script loaded\r";
});
Office.onReady(function (info) {
  document.getElementById("logger").innerHTML += "office ready: " + JSON.stringify(info) + "\r";
  $(document).ready(function () {
    document.getElementById("logger").innerHTML += "document ready\r";
    if (info.host === Office.HostType.Outlook) {
      document.getElementById("logger").innerHTML += "outlook\r";
      document.getElementById("submit").onclick = run;
      loadProps();
    }
  });
});

function run() {
  document.getElementById("logger").innerHTML += "run\r";
  fillData(gItem);
}

function loadProps() {
  document.getElementById("logger").innerHTML += "load props\r";
  var item = Office.context.mailbox.item;
  document.getElementById("logger").innerHTML += JSON.stringify(item) + "\r";
  item.body.getAsync("text", { asyncContext: "callback" }, function (data) {
    gItem = {
      start: item.start,
      end: item.end,
      location: item.location,
      subject: item.subject,
      optionalAttendees: item.optionalAttendees,
      requiredAttendees: item.requiredAttendees,
      body: data.value,
      organizer: item.organizer
    };
    renderForm(gItem);
  });
}

function renderForm(item) {
  $('#start').val(item.start.format('yyyy-MM-dd'));
  $('#end').text(item.end);
  $('#location').html(item.location);
  $('#normalizedSubject').text(item.subject);
  $('#optionalAttendees').html(buildEmailAddressesString(item.optionalAttendees));
  $('#requiredAttendees').html(buildEmailAddressesString(item.requiredAttendees));
  $('#body').html(item.body);
}

function fillData(item) {
  document.getElementById("logger").innerHTML += "fill\r";
  var message = 'pageId=53811457&f=meetingCollector&title01=' + item.subject +
    '&beginTm=' + item.start.format('dd.MM.yyyy HH:mm') +
    '&endTm=' + item.end.format('dd.MM.yyyy HH:mm') +
    '&obligMember=' + item.requiredAttendees.map(function (address) { return address.emailAddress; }) +
    '&optionalMember=' + item.optionalAttendees.map(function (address) { return address.emailAddress; }) +
    '&place=' + item.location +
    '&agenda=' + item.body +
    '&type=OutlookConfluence' +
    '&authorMeeting=' + item.organizer.emailAddress;

  sendMessage(message);

  document.getElementById("logger").innerHTML += message + "\r";
}

function sendMessage(message) {
  document.getElementById("logger").innerHTML += "send message\r";
  $.ajax({
    url: 'https://confluence.beeline.kz/ajax/confiforms/rest/save.action',
    type: 'POST',
    headers: { "Authorization": "Basic " + btoa("tech_outlook_mom:~F4B?#?Z") },
    contentType: "application/x-www-form-urlencoded;",
    data: message
  }).done(function (data) {
    document.getElementById("logger").innerHTML += "send message done\r";
    try {
      var jsonData = JSON.parse(data);
      document.getElementById("logger").innerHTML += JSON.stringify(jsonData) + "\r";
      getItemUrl(jsonData.id);
    }
    catch (ex) {
      document.getElementById("logger").innerHTML += ex.message + "\r";
      return;
    }

  }).fail(function (error) {
    document.getElementById("logger").innerHTML += "send message error: " + error + "\r";
  });
}

function getItemUrl(id) {
  document.getElementById("logger").innerHTML += "get item url: " + id + "\r";
  $.ajax({
    url: 'https://confluence.beeline.kz/ajax/confiforms/rest/filter.action',
    type: 'GET',
    headers: { "Authorization": "Basic " + btoa("tech_outlook_mom:~F4B?#?Z") },
    contentType: "application/x-www-form-urlencoded;",
    data: 'pageId=53811457&f=meetingCollector&q=id:' + id
  }).done(function (data) {
    document.getElementById("logger").innerHTML += "get item url done\r";
    var pId = data.list.entry[0].fields.meetingLink;
    window.open('https://confluence.beeline.kz/pages/viewpage.action?pageId=' + pId, '_blank');
  }).fail(function (ctx, status, error) {
    document.getElementById("logger").innerHTML += JSON.stringify(ctx) + "\r";
  });
}

// Форматировать объект EmailAddressDetails как
// Имя Фамилия <emailaddress>
function buildEmailAddressString(address) {
  return "<a href='" + address.emailAddress + "'>" + address.displayName + "</a>";
}

// Взять массив объектов AttachmentDetails и
// создать список форматированных строк, разделенных разрывом строки
function buildEmailAddressesString(addresses) {
  if (addresses && addresses.length > 0) {
    var returnString = "";

    for (var i = 0; i < addresses.length; i++) {
      if (i > 0) {
        returnString = returnString + "<br />";
      }
      returnString = returnString + buildEmailAddressString(addresses[i]);
    }

    return returnString;
  }

  return "None";
}