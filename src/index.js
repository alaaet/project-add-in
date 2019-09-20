/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
$.getScript("../models/task.js", function() {});

$(document).ready(() => {
  $("#run").click(run);
});

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
  $("#sideload-msg").hide();
  $("#app-body").show();
};

function run() {
  var tasks = [];

  getMaxTaskIndex().then(function(maxIndex) {
    for (let index = 1; index < maxIndex + 1; index++) {
      var name, duration, start, finish;
      getTaskGuid(index)
        .then(function(taskGuid) {
          console.log("task GUID:" + taskGuid);

          getTaskAttribute(taskGuid, Office.ProjectTaskFields.Name).then(
            function(value) {
              name = value;
            }
          );
          getTaskAttribute(taskGuid, Office.ProjectTaskFields.Duration).then(
            function(value) {
              duration = value;
            }
          );
          getTaskAttribute(taskGuid, Office.ProjectTaskFields.Start).then(
            function(value) {
              start = value;
            }
          );
          getTaskAttribute(taskGuid, Office.ProjectTaskFields.Finish).then(
            function(value) {
              finish = value;
            }
          );
        })
        .then(function() {
          var task = new Task(index, name, duration, start, finish);
          tasks.push(task);
          if (index == maxIndex) {
            console.log(JSON.stringify(tasks, null, 2));
          }
        });
    }
  });
}

// Get the GUID of a task.
function getTaskGuid(index) {
  var defer = $.Deferred();
  Office.context.document.getTaskByIndexAsync(index, function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      onError(result.error);
    } else {
      defer.resolve(result.value);
    }
  });
  return defer.promise();
}

// Get the maximum index of the tasks for the current project.
function getMaxTaskIndex() {
  var defer = $.Deferred();
  Office.context.document.getMaxTaskIndexAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      onError(result.error);
    } else {
      defer.resolve(result.value);
    }
  });
  return defer.promise();
}

// Get the attribute value for a specific task.
function getTaskAttribute(taskGuid, targetField) {
  var defer = $.Deferred();
  Office.context.document.getTaskFieldAsync(taskGuid, targetField, function(
    result
  ) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      onError(result.error);
    } else {
      //console.log(result.value.fieldValue);
      defer.resolve(result.value.fieldValue);
    }
  });
  return defer.promise();
}

function onError(error){
  console.log("ERROR: "+error)
}