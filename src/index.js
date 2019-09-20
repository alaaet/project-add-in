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
var tasks = [];
function run() {
  getMaxTaskIndex().then(function(maxIndex) {
    console.log("Max index= " + maxIndex);
    /*for (let index = 0; index <= maxIndex; index++) {
      var guid, name, duration, start, finish, resourceNames;
      getTaskGuid(index)
        .then(function(taskGuid) {
          console.log(index + " -task GUID:" + taskGuid);
          guid = taskGuid;
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
          getTaskAttribute(
            taskGuid,
            Office.ProjectTaskFields.ResourceNames
          ).then(function(value) {
            resourceNames = value;
          });
        })
        .then(function() {
          var task = new Task(
            index,
            guid,
            name,
            duration,
            start,
            finish,
            resourceNames
          );
          tasks.push(task);
          if (index == maxIndex) {
            // SEND THE DATA OF ALL TASKS TO THE API
            submitData(tasks);
          }
        });
    }*/

    for (let index = 0; index <= maxIndex; index++) {
      getTaskObject(index, maxIndex);
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

// Get all the basic attributes values for a specific task.
function getTaskObject(index, maxIndex) {
  var guid, name, duration, start, finish, resourceNames;
  getTaskGuid(index).then(function(taskGuid) {
    guid = taskGuid;
    getTaskAttribute(taskGuid, Office.ProjectTaskFields.Name).then(function(
      value
    ) {
      name = value;
      getTaskAttribute(taskGuid, Office.ProjectTaskFields.Duration).then(
        function(value) {
          duration = value;
          getTaskAttribute(taskGuid, Office.ProjectTaskFields.Start).then(
            function(value) {
              start = value;
              getTaskAttribute(taskGuid, Office.ProjectTaskFields.Finish).then(
                function(value) {
                  finish = value;
                  getTaskAttribute(
                    taskGuid,
                    Office.ProjectTaskFields.ResourceNames
                  ).then(function(value) {
                    resourceNames = value;
                    var task = new Task(
                      index,
                      guid,
                      name,
                      duration,
                      start,
                      finish,
                      resourceNames
                    );
                    //console.log(JSON.stringify(task));
                    tasks.push(task);
                    if (index == maxIndex) {
                      submitData(tasks);
                    }
                  });
                }
              );
            }
          );
        }
      );
    });
  });
}

// Send data to API
function submitData(tasks) {
  $.each(tasks, function(index, task) {
    console.dir(JSON.stringify(task));
  });

  // @Anand: here you can make the request to the API
}

// SUPPORT FUNCTIONS
function onError(error) {
  console.log("ERROR: " + error);
}
