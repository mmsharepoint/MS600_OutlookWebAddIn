(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.onReady(function () {
    $(document).ready(function () {
      var element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();
      var savBtn = document.getElementById('saveMail');
      savBtn.addEventListener('click', saveMimeMail);
      loadFileTypes();
    });
  });
  
  async function loadFileTypes() {
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken();
    $.ajax({
      url: '/api/Web/FileTypes',
      accepts: 'application/json',
      headers: {
        "Authorization": "Bearer " + bootstrapToken // Used here to pass authorization in WebController
      }
    })
      .done((response) => {
        var select = document.getElementById('fileTypes');
        response.forEach((val) => {
          var opt = document.createElement("option");
          opt.value = val;
          opt.text = val;
          select.options.add(opt);
        });
      });
    renderAttachments(bootstrapToken);
  }

  async function saveMimeMail() {
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken();
    const mailID = Office.context.mailbox.item.itemId;
    const restMailID = Office.context.mailbox.convertToRestId(mailID, Office.MailboxEnums.RestVersion.v2_0);
    const requestBody = { MessageID: restMailID };
    $.ajax({
      type: "POST",
      url: '/api/Web/StoreMimeMessage',
      accepts: 'application/json',
      headers: {
        "Authorization": "Bearer " + bootstrapToken // Used here to pass authorization in WebController
      },
      data: JSON.stringify(requestBody),
      contentType: "application/json; charset=utf-8"
    }).done(function (data) {
      console.log(data);
    }).fail(function (error) {
      console.log(error);
    });
  }

  async function renderAttachments(bootstrapToken) {
    const mailID = Office.context.mailbox.item.itemId;
    const restMailID = Office.context.mailbox.convertToRestId(mailID, Office.MailboxEnums.RestVersion.v2_0);
    const requestBody = { MessageID: restMailID };
    $.ajax({
      url: '/api/Web/GetAttachments',
      type: 'POST',
      accepts: 'application/json',
      headers: {
        "Authorization": "Bearer " + bootstrapToken // Used here to pass authorization in WebController
      },
      data: JSON.stringify(requestBody),
      contentType: "application/json; charset=utf-8"
    }).done(function (data) {
      console.log(data);
      var list = document.getElementById('attachmentsList');
      data.forEach((doc) => {
        var listItem = document.createElement('li');
        listItem.innerHTML = '<input type="checkbox" data-docID="' + doc.id + '" data-docName="' + doc.name + '" /> ' + doc.name;
        list.appendChild(listItem);
      });
      var savBtn = document.getElementById('saveAttachments');
      savBtn.addEventListener('click', saveAttachments);
    }).fail(function (error) {
      console.log(error);
    });
  }

  async function saveAttachments() {
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken();
    var attachments = document.querySelectorAll('.attachmentsList input[type="checkbox"]:checked');
    var attArr = Array.from(attachments);
    var selectedDocs = [];
    attArr.forEach((sel) => {
      console.log(sel.getAttribute('data-docID'));
      selectedDocs.push({ id: sel.getAttribute('data-docID'), filename: sel.getAttribute('data-docName') });
    });
    const mailID = Office.context.mailbox.item.itemId;
    const restMailID = Office.context.mailbox.convertToRestId(mailID, Office.MailboxEnums.RestVersion.v2_0);
    const requestBody = { Attachments: selectedDocs, MessageID: restMailID };
    $.ajax({
      url: '/api/Web/SaveAttachments',
      type: 'POST',
      accepts: 'application/json',
      headers: {
        "Authorization": "Bearer " + bootstrapToken // Used here to pass authorization in WebController
      },
      data: JSON.stringify(requestBody),
      contentType: "application/json; charset=utf-8"
    }).done(function (data) {
      console.log(data);      
    }).fail(function (error) {
      console.log(error);
    });
  }
  // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }

      return returnString;
    }

    return "None";
  }

  // Format an EmailAddressDetails object as
  // GivenName Surname <emailaddress>
  function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }

  // Take an array of EmailAddressDetails objects and
  // build a list of formatted strings, separated by a line-break
  function buildEmailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";

      for (var i = 0; i < addresses.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + buildEmailAddressString(addresses[i]);
      }

      return returnString;
    }

    return "None";
  }

  // Load properties from the Item base object, then load the
  // message-specific properties.
  function loadProps() {
    var item = Office.context.mailbox.item;

    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    $('#message-props').show();

    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.from));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
    $('#to').html(buildEmailAddressesString(item.to));
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();