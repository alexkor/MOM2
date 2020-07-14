/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
$(document).ready(function () {
  document.getElementById("logger").innerHTML = "script loaded\r";
});
Office.onReady(function (info) {
  document.getElementById("logger").innerHTML = "office ready\r";
  $(document).ready(function () {
    document.getElementById("logger").innerHTML = "document ready\r";
    if (info.host === Office.HostType.Outlook) {
      document.getElementById("logger").innerHTML += "outlook\r";
      document.getElementById("run").onclick = run;
    }
  });
});

function run() {
  document.getElementById("logger").innerHTML += "run\r";
  loadProps();
}

function loadProps() {
  document.getElementById("logger").innerHTML += "load props\r";
  var item = Office.context.mailbox.item;
  item.body.getAsync("text", { asyncContext: "callback" }, function (data) {
    fillData(item, data.value);
  });
}

function fillData(item, body) {
  document.getElementById("logger").innerHTML += "fill\r";
  var message = 'pageId=53811457&f=meetingCollector&title01=' + item.subject +
    '&beginTm=' + item.start.format('dd.MM.yyyy HH:mm') +
    '&endTm=' + item.end.format('dd.MM.yyyy HH:mm') +
    '&obligMember=' + item.requiredAttendees.map(function (address) { return address.emailAddress; }) +
    '&optionalMember=' + item.optionalAttendees.map(function (address) { return address.emailAddress; }) +
    '&place=' + item.location +
    '&agenda=' + body +
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